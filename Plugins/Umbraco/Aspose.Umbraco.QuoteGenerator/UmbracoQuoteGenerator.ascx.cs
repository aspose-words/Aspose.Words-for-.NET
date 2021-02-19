using System;
using System.Web;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.MailMerging;
using ImageFormat = System.Drawing.Imaging.ImageFormat;

namespace Aspose.UmbracoQuoteGenerator
{
    public partial class UmbracoQuoteGenerator : System.Web.UI.UserControl
    {
        #region Page load and events

        // Page load event.
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                // Verify that the page is not post back, so we can setup default page view.
                if (!Page.IsPostBack)
                {
                    // Calling this function to create default rows when page initially loaded.
                    PopulateProductsGrid(int.Parse((txtAddProductRows.Text.Trim().Equals("") == false ? txtAddProductRows.Text.Trim() : "3")));
                    txtDocDate.Text = DateTime.Now.ToLongDateString();
                    txtDocNo.Text = DateTime.Now.ToShortDateString() + "-001";
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        // Adding rows in invoice products gridview.
        protected void btnAddProducts_Click(object sender, EventArgs e)
        {
            try
            {
                // Verify that the textbox is not empty.
                if (!txtAddProductRows.Text.Trim().Equals(""))
                {
                    // Populate product empty rows as per user input for rows.
                    PopulateProductsGrid(int.Parse(txtAddProductRows.Text.Trim()));
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        // Button even to generate invoice in PDF file.
        protected void btnGeneratePDF_Click(object sender, EventArgs e)
        {
            try
            {
                lblMessage.Text = "";
                string TargetPathFileSave = "";

                // Verify logo file is selected by user to upload.
                if (fuCompanyLogo.HasFile)
                {
                    int imgSize = fuCompanyLogo.PostedFile.ContentLength;

                    string ext = System.IO.Path.GetExtension(this.fuCompanyLogo.PostedFile.FileName);
                    if (ext.ToUpper().Trim() != ".JPG" && ext.ToUpper() != ".PNG" && ext.ToUpper() != ".GIF" && ext.ToUpper() != ".JPEG")
                    {
                        lblMessage.Text = "Please choose only .jpg, .png and .gif image types";
                        return;
                    }

                    if (imgSize > 1048576)
                    {
                        lblMessage.Text = "Maximum image file size 1 MB";
                        return;
                    }

                    // Verify and secure your upload that only allow image files and no security risks attached.
                    System.Drawing.Image image = System.Drawing.Image.FromStream(fuCompanyLogo.FileContent);

                    if (!new[] { ImageFormat.Gif.Guid, ImageFormat.Jpeg.Guid, ImageFormat.Bmp.Guid, ImageFormat.Png.Guid, ImageFormat.Icon.Guid }.Contains(image.RawFormat.Guid))
                        throw new ArgumentException("Invalid image file type");

                    // Base directory path to upload image.
                    TargetPathFileSave = Server.MapPath(GetDataDir_LogoImages());

                    // Apply a GUID to create unique file names.
                    Guid nimgGUID = Guid.NewGuid();
                    TargetPathFileSave = TargetPathFileSave + nimgGUID.ToString().Trim() + fuCompanyLogo.FileName.Substring(fuCompanyLogo.FileName.LastIndexOf('.')).ToLower();

                    // Upload file to the server.
                    fuCompanyLogo.PostedFile.SaveAs(TargetPathFileSave);
                }
                else
                {
                    // If no file selected, then the user should provide the company name.
                    if (txtCompanyName.Text.Trim().Equals(""))
                    {
                        // If no file and company name is entered, notify user and stop process.
                        lblMessage.Text = "please select file to upload.";
                        return;
                    }
                }

                // Generating PDF for user input using template document.
                MergeWithWordTemplate(Server.MapPath(GetDataDir_Templates()), Server.MapPath(GetDataDir_OutputDocs()), TargetPathFileSave);

            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        // Button event to clear form fields.
        protected void btnClearForm_Click(object sender, EventArgs e)
        {
            try
            {
                // In this demo example one page by redirecting to same page will reset all fields to its initial state.
                Response.Redirect(Request.Url.ToString());
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        // Products gridview row data bound event to populate VAT dropdown list.
        protected void grdInvoiceProducts_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            try
            {
                if (e.Row.RowType == DataControlRowType.DataRow)
                {
                    // Getting object of VAT dropdown list for each row.
                    DropDownList ddlProductVAT = (DropDownList)e.Row.FindControl("ddlProductVAT");

                    if (ddlProductVAT != null)
                    {
                        QuoteGenerator.PopulateVATDropdownList(ref ddlProductVAT, this.Session);
                    }
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        #endregion

        #region Private methods

        // Add rows in invoice products gridview.
        private void PopulateProductsGrid(int addRows)
        {
            try
            {
                // Get dataset and datable object from cache.
                DataSet data = QuoteGenerator.GetDataSetForGridView(Session);
                if (data != null)
                {
                    // Remove all rows in collection.
                    data.Tables[0].Rows.Clear();

                    for (int indx = 1; indx <= addRows; indx++)
                    {
                        // Add the temp data row to the tables for each row.
                        data.Tables[0].Rows.Add(indx, "", 0.0, 1, 0.0, 0.0, 0.0, 0.0);
                    }

                    grdInvoiceProducts.DataSource = data;
                    grdInvoiceProducts.DataBind();
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        private void MergeWithWordTemplate(string templatePath, string outputPath, string imagePath)
        {
            try
            {
                lblMessage.Text = "";

                // Check for license and apply if exists.
                string licenseFile = Server.MapPath("~/App_Data/Aspose.Words.lic");
                if (File.Exists(licenseFile))
                {
                    License license = new License();
                    license.SetLicense(licenseFile);
                }

                Document doc = QuoteGenerator.GetUnmergedTemplateObject(templatePath + "MailMerge_Template.doc", Session);
                if (doc != null)
                {
                    // Fill the fields in the document with user data.
                    DataSet data = QuoteGenerator.GetDataSetForGridView(Session);
                    if (data != null)
                    {
                        decimal grandTotalAllItemsAmount = 0;

                        if (grdInvoiceProducts.Rows.Count > 0)
                        {
                            data.Tables[0].Rows.Clear();

                            System.Web.UI.WebControls.TextBox txtProductDescription;
                            System.Web.UI.WebControls.TextBox txtProductPrice;
                            System.Web.UI.WebControls.TextBox txtProductQuantity;
                            DropDownList ddlProductVAT;

                            foreach (GridViewRow gr in grdInvoiceProducts.Rows)
                            {
                                // Find control in each gridview row.
                                txtProductDescription = (System.Web.UI.WebControls.TextBox)gr.FindControl("txtProductDescription");
                                txtProductPrice = (System.Web.UI.WebControls.TextBox)gr.FindControl("txtProductPrice");
                                txtProductQuantity = (System.Web.UI.WebControls.TextBox)gr.FindControl("txtProductQuantity");
                                ddlProductVAT = (DropDownList)gr.FindControl("ddlProductVAT");

                                // Verify the found controls should not be null.
                                if (txtProductDescription != null && txtProductPrice != null && txtProductQuantity != null && ddlProductVAT != null)
                                {
                                    // Verify the found controls should not be empty.
                                    if (txtProductDescription.Text.Trim() != "" && txtProductPrice.Text.Trim() != "" && txtProductQuantity.Text.Trim() != "" && ddlProductVAT.Items.Count > 0)
                                    {
                                        // Actual amount = price * quantity.
                                        decimal itemTotalBeforeVat = decimal.Parse(txtProductPrice.Text.Trim()) * decimal.Parse(txtProductQuantity.Text.Trim());

                                        // VAT amount = (actual amount * VAT) / 100 .
                                        decimal itemTotalVatAmount = (itemTotalBeforeVat * decimal.Parse(ddlProductVAT.SelectedItem.Value.Trim())) / 100;

                                        // Total amount including VAT.
                                        decimal itemTotalAmount = itemTotalBeforeVat + itemTotalVatAmount;
                                        grandTotalAllItemsAmount += itemTotalAmount;

                                        // Add the temp data row to the tables for each row.
                                        data.Tables[0].Rows.Add(gr.Cells[0].Text, txtProductDescription.Text.Trim(), decimal.Parse(txtProductPrice.Text.Trim()), decimal.Parse(txtProductQuantity.Text.Trim()), itemTotalBeforeVat, decimal.Parse(ddlProductVAT.SelectedItem.Value), itemTotalVatAmount, itemTotalAmount);
                                    }
                                }
                            }
                        }

                        if (imagePath != "")
                        {
                            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
                            if (shape != null)
                            {
                                shape.ImageData.ImageBytes = File.ReadAllBytes(imagePath);
                            }
                        }
                        else
                        {
                            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
                            if (shape != null)
                            {
                                shape.Remove();
                            }
                        }

                        // Update fields using an Aspose.Words Mail Merge.
                        doc.MailMerge.Execute(
                            new string[] { "CompanyName", "CompanyAddress", "CompanyZipState", "CompanyCountry", "CustomerName", "CustomerAddress", "CustomerZipState", "CustomerCountry", "InvoiceTotalAmount", "DocCaption", "DocDate", "DocNo", "DocDescription", "DocTC" },
                            new object[] { txtCompanyName.Text.Trim(), txtCompanyAddress.Text.Trim(), txtCompanyStateZip.Text.Trim(), txtCompanyCountry.Text.Trim(), txtCustomerName.Text.Trim(), txtCustomerAddress.Text.Trim(), txtCustomerStateZip.Text.Trim(), txtCustomerCountry.Text.Trim(), grandTotalAllItemsAmount, txtDocCaption.Text.Trim(), txtDocDate.Text.Trim(), txtDocNo.Text.Trim(), txtDescription.Text.Trim(), txtTC.Text.Trim() });

                        doc.MailMerge.ExecuteWithRegions(data);

                        // Remove unused fields in template.
                        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs | MailMergeCleanupOptions.RemoveContainingFields | MailMergeCleanupOptions.RemoveUnusedFields;

                        // Updated document layout, to be cached and re-use.
                        doc.UpdatePageLayout();

                        // Saves the document to disk.
                        string fname = System.Guid.NewGuid().ToString() + "." + QuoteGenerator.GetSaveFormat(ExportTypeDropDown.SelectedValue);
                        doc.Save(outputPath + fname);
                        Response.Clear();
                        Response.Buffer = true;
                        Response.AddHeader("content-disposition", "attachment;filename=ExportedFile_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + "_" + DateTime.Now.Millisecond.ToString() + "." + QuoteGenerator.GetSaveFormat(ExportTypeDropDown.SelectedValue));
                        Response.Charset = "";
                        Response.ContentType = "application/pdf";
                        Response.Cache.SetCacheability(HttpCacheability.NoCache);

                        Response.ContentType = "Application/" + QuoteGenerator.GetSaveFormat(ExportTypeDropDown.SelectedValue);

                        // Get the physical path to the file.
                        string FilePath = MapPath(GetDataDir_OutputDocs() + fname);

                        // Write the file directly to the HTTP content output stream.
                        Response.WriteFile(FilePath);
                        Response.Flush();

                        // Delete file as its already in stream and available for user to download/save/view.
                        FileInfo file = new FileInfo(FilePath);
                        if (file.Exists)
                        {
                            file.Delete();
                        }
                        file = new FileInfo(imagePath);
                        if (file.Exists)
                        {
                            file.Delete();
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
                Response.Clear();
                Response.Flush();
            }
        }

        #endregion

        #region Folder Paths

        // Path to merge templates folder.
        private string GetDataDir_Templates()
        {
            try
            {
                if (!System.IO.Directory.Exists(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/Templates/")))
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/Templates/"));

                return "~/UserControls/Aspose.UmbracoQuoteGenerator/Templates/";
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
                throw exc;
            }
        }

        // Path to logo images folder.
        private string GetDataDir_LogoImages()
        {
            try
            {
                if (!System.IO.Directory.Exists(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/UploadedImages/")))
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/UploadedImages/"));

                return "~/UserControls/Aspose.UmbracoQuoteGenerator/UploadedImages/";
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
                throw exc;
            }
        }

        // Path to output documents folder.
        private string GetDataDir_OutputDocs()
        {
            try
            { 
                if (!System.IO.Directory.Exists(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/OutputDocs/")))
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/OutputDocs/"));

                return "~/UserControls/Aspose.UmbracoQuoteGenerator/OutputDocs/";
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
                throw exc;
            }
        }

        // Path to output documents folder.
        private string GetDataDir_License()
        {
            try
            {
                if (!System.IO.Directory.Exists(Server.MapPath("~/bin/")))
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/bin/"));

                return "~/bin/";
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
                throw exc;
            }
        }

        #endregion
    }
}