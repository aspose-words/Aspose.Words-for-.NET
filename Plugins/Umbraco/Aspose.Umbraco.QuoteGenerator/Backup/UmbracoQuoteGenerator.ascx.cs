using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Collections;
using Aspose.Words.Layout;
using Aspose.Words.Saving;
using Aspose.Words.MailMerging;

namespace Aspose.UmbracoQuoteGenerator
{
    public partial class UmbracoQuoteGenerator : System.Web.UI.UserControl
    {
        #region Page load and events

        // page load event
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                // verify page is not post back, so we can setup default page view.
                if (!Page.IsPostBack)
                {
                    // calling this function to create default rows when page initialy loaded.
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

        // adding rows in invoice products gridview.
        protected void btnAddProducts_Click(object sender, EventArgs e)
        {
            try
            {
                // varify textbox is not empty
                if (!txtAddProductRows.Text.Trim().Equals(""))
                {
                    // populating product empty rows as per user input for rows
                    PopulateProductsGrid(int.Parse(txtAddProductRows.Text.Trim()));
                }
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        // button even to generate invoice in PDF file
        protected void btnGeneratePDF_Click(object sender, EventArgs e)
        {
            try
            {
                lblMessage.Text = "";
                string TargetPathFileSave = "";

                // verify logo file is selected by user to upload
                if (fuCompanyLogo.HasFile)
                {
                    int imgSize = fuCompanyLogo.PostedFile.ContentLength;

                    string ext = System.IO.Path.GetExtension(this.fuCompanyLogo.PostedFile.FileName);
                    if (ext.ToUpper().Trim() != ".JPG" && ext.ToUpper() != ".PNG" && ext.ToUpper() != ".GIF" && ext.ToUpper() != ".JPEG")
                    {
                        lblMessage.Text = "Please choose only .jpg, .png and .gif image types";
                        return;
                    }
                    else
                    {
                        if (imgSize > 1048576)
                        {
                            lblMessage.Text = "Maximum image file size 1 MB";
                            return;
                        }
                    }
                    // Verify and secure your upload that only allow image files and no security risks attached
                    System.Drawing.Image image = System.Drawing.Image.FromStream(fuCompanyLogo.FileContent);
                    string FormetType = string.Empty;
                    if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Gif.Guid)
                        FormetType = "GIF";
                    else if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Jpeg.Guid)
                        FormetType = "JPG";
                    else if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Bmp.Guid)
                        FormetType = "BMP";
                    else if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Png.Guid)
                        FormetType = "PNG";
                    else if (image.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Icon.Guid)
                        FormetType = "ICO";
                    else
                        throw new System.ArgumentException("Invalid File Type");

                    // base directory path to upload image
                    TargetPathFileSave = Server.MapPath(GetDataDir_LogoImages());

                    // use GUID to distinct each file name
                    Guid nimgGUID = Guid.NewGuid();
                    TargetPathFileSave = TargetPathFileSave + nimgGUID.ToString().Trim() + fuCompanyLogo.FileName.Substring(fuCompanyLogo.FileName.LastIndexOf('.')).ToLower();

                    // upload file to server
                    fuCompanyLogo.PostedFile.SaveAs(TargetPathFileSave);
                }
                else
                {
                    // if no file selected then company name should be provided by user
                    if (txtCompanyName.Text.Trim().Equals(""))
                    {
                        // in case no file and company name provided then notify user and stop process
                        lblMessage.Text = "please select file to upload.";
                        return;
                    }
                }

                // generating PDF for user input using template document
                MergeWithWordTemplate(Server.MapPath(GetDataDir_Templates()), Server.MapPath(GetDataDir_OutputDocs()), TargetPathFileSave);

            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        // button event to clear form fields
        protected void btnClearForm_Click(object sender, EventArgs e)
        {
            try
            {
                // in this demo example one page by redirecting to same page will reset all fields to its initial state
                Response.Redirect(Request.Url.ToString());
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
            }
        }

        // products gridview row data bound event to populate VAT dropdown list
        protected void grdInvoiceProducts_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            try
            {
                if (e.Row.RowType == DataControlRowType.DataRow)
                {
                    // getting object of VAT dropdown list for each row
                    DropDownList ddlProductVAT = (DropDownList)e.Row.FindControl("ddlProductVAT");

                    // verify dropdown list object is not null
                    if (ddlProductVAT != null)
                    {
                        // call populate VAT dropdown list
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

        // adding rows in invoice products gridview
        private void PopulateProductsGrid(int addRows)
        {
            try
            {
                // get dataset and datable object from cache.
                DataSet data = QuoteGenerator.GetDataSetForGridView(Session);
                if (data != null)
                {
                    // removing all rows in collection
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

                // Check for license and apply if exists
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
                        decimal itemtotalBeforeVAT = 0;
                        decimal itemtotalVATAmount = 0;
                        decimal itemtotalAmount = 0;
                        decimal grandTotalAllItemsAmount = 0;

                        if (grdInvoiceProducts.Rows.Count > 0)
                        {
                            // removing all rows in collection
                            data.Tables[0].Rows.Clear();

                            System.Web.UI.WebControls.TextBox txtProductDescription;
                            System.Web.UI.WebControls.TextBox txtProductPrice;
                            System.Web.UI.WebControls.TextBox txtProductQuantity;
                            DropDownList ddlProductVAT;

                            foreach (GridViewRow gr in grdInvoiceProducts.Rows)
                            {
                                // find control in each gridview rows
                                txtProductDescription = (System.Web.UI.WebControls.TextBox)gr.FindControl("txtProductDescription");
                                txtProductPrice = (System.Web.UI.WebControls.TextBox)gr.FindControl("txtProductPrice");
                                txtProductQuantity = (System.Web.UI.WebControls.TextBox)gr.FindControl("txtProductQuantity");
                                ddlProductVAT = (DropDownList)gr.FindControl("ddlProductVAT");

                                // varify the found controls should not be null
                                if (txtProductDescription != null && txtProductPrice != null && txtProductQuantity != null && ddlProductVAT != null)
                                {
                                    // varify the found controls should not be empty
                                    if (txtProductDescription.Text.Trim() != "" && txtProductPrice.Text.Trim() != "" && txtProductQuantity.Text.Trim() != "" && ddlProductVAT.Items.Count > 0)
                                    {
                                        // actual amount price X quantity
                                        itemtotalBeforeVAT = (decimal.Parse(txtProductPrice.Text.Trim()) * decimal.Parse(txtProductQuantity.Text.Trim()));

                                        // VAT amount = (actual X VAT)/100
                                        itemtotalVATAmount = ((itemtotalBeforeVAT * decimal.Parse(ddlProductVAT.SelectedItem.Value.Trim())) / 100);

                                        // Total amount including VAT
                                        itemtotalAmount = itemtotalBeforeVAT + itemtotalVATAmount;
                                        grandTotalAllItemsAmount += itemtotalAmount;

                                        // Add the temp data row to the tables for each row.
                                        data.Tables[0].Rows.Add(gr.Cells[0].Text, txtProductDescription.Text.Trim(), decimal.Parse(txtProductPrice.Text.Trim()), decimal.Parse(txtProductQuantity.Text.Trim()), itemtotalBeforeVAT, decimal.Parse(ddlProductVAT.SelectedItem.Value), itemtotalVATAmount, itemtotalAmount);
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
                        // updating fix fields using simple aspose mail merge
                        doc.MailMerge.Execute(
                            new string[] { "CompanyName", "CompanyAddress", "CompanyZipState", "CompanyCountry", "CustomerName", "CustomerAddress", "CustomerZipState", "CustomerCountry", "InvoiceTotalAmount", "DocCaption", "DocDate", "DocNo", "DocDescription", "DocTC" },
                            new object[] { txtCompanyName.Text.Trim(), txtCompanyAddress.Text.Trim(), txtCompanyStateZip.Text.Trim(), txtCompanyCountry.Text.Trim(), txtCustomerName.Text.Trim(), txtCustomerAddress.Text.Trim(), txtCustomerStateZip.Text.Trim(), txtCustomerCountry.Text.Trim(), grandTotalAllItemsAmount, txtDocCaption.Text.Trim(), txtDocDate.Text.Trim(), txtDocNo.Text.Trim(), txtDescription.Text.Trim(), txtTC.Text.Trim() });

                        doc.MailMerge.ExecuteWithRegions(data);

                        // removing unused fields in template
                        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs | MailMergeCleanupOptions.RemoveContainingFields | MailMergeCleanupOptions.RemoveUnusedFields;

                        // updating document layout, to be cached and re-use
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
                        //Get the physical path to the file.
                        string FilePath = MapPath(GetDataDir_OutputDocs() + fname);

                        //Write the file directly to the HTTP content output stream.
                        Response.WriteFile(FilePath);
                        Response.Flush();

                        // delete file as its already in stream and available for user to download/save/view.
                        FileInfo file = new FileInfo(FilePath);
                        if (file.Exists)//check file exsit or not
                        {
                            file.Delete();
                        }
                        file = new FileInfo(imagePath);
                        if (file.Exists)//check file exsit or not
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

        // path to merge templates folder
        private string GetDataDir_Templates()
        {
            try
            {
                // check if directory exist
                if (!System.IO.Directory.Exists(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/Templates/")))
                {
                    // create directory if missing
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/Templates/"));
                }
                return "~/UserControls/Aspose.UmbracoQuoteGenerator/Templates/";
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
                throw exc;
            }
        }

        // path to logo images folder
        private string GetDataDir_LogoImages()
        {
            try
            {
                // check if directory exist
                if (!System.IO.Directory.Exists(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/UploadedImages/")))
                {
                    // create directory if missing
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/UploadedImages/"));
                }
                return "~/UserControls/Aspose.UmbracoQuoteGenerator/UploadedImages/";
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
                throw exc;
            }
        }

        // path to output documents folder
        private string GetDataDir_OutputDocs()
        {
            try
            {
                // check if directory exist
                if (!System.IO.Directory.Exists(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/OutputDocs/")))
                {
                    // create directory if missing
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/UserControls/Aspose.UmbracoQuoteGenerator/OutputDocs/"));
                }
                return "~/UserControls/Aspose.UmbracoQuoteGenerator/OutputDocs/";
            }
            catch (Exception exc)
            {
                lblMessage.Text = exc.Message;
                throw exc;
            }
        }

        // path to output documents folder
        private string GetDataDir_License()
        {
            try
            {
                // check if directory exist
                if (!System.IO.Directory.Exists(Server.MapPath("~/bin/")))
                {
                    // create directory if missing
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/bin/"));
                }
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