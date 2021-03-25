using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI.WebControls;
using umbraco.cms.businesslogic.member;
using Aspose.Words;
using System.Drawing;
using System.IO;
using Aspose.Words.Tables;

namespace Aspose.UmbracoMemberExportToWord
{
    public partial class AsposeMemberExport : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                ErrorLabel.Visible = false;
                if (!Page.IsPostBack)
                    LoadMembers();
            }
            catch (Exception ex)
            {
                ErrorLabel.Text = ex.ToString();
                ErrorLabel.Visible = true;
            }
        }

        void LoadMembers()
        {
            IEnumerable<Member> listMembers = Member.GetAllAsList();
            UmbracoMembersGridView.DataSource = listMembers;
            UmbracoMembersGridView.DataBind();

            if (UmbracoMembersGridView.Rows.Count > 0)
            {
                UmbracoMembersGridView.UseAccessibleHeader = true;
                UmbracoMembersGridView.HeaderRow.TableSection = TableRowSection.TableHeader;
            }
        }

        protected void ExportButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Check for an Aspose.Words license file in the local file system, and then apply it if it exists.
                string licenseFile = Server.MapPath("~/App_Data/Aspose.Words.lic");
                if (File.Exists(licenseFile))
                {
                    License license = new License();
                    license.SetLicense(licenseFile);
                }

                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                Aspose.Words.Tables.Table table = builder.StartTable();

                // Make the header row.
                builder.InsertCell();

                // Set the left indent for the table. Table wide formatting must be applied after 
                // at least one row is present in the table.
                table.LeftIndent = 20.0;

                // Set height and define the height rule for the header row.
                builder.RowFormat.Height = 40.0;
                builder.RowFormat.HeightRule = HeightRule.AtLeast;

                // Some special features for the header row.
                builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Font.Size = 10;
                builder.Font.Name = "Arial";
                builder.Font.Bold = true;
                builder.Write("Name");

                // We don't need to specify the width of this cell because it's inherited from the previous cell.
                builder.InsertCell();
                builder.Write("LoginName");

                builder.InsertCell();
                builder.Write("Email");

                // We don't need to specify the width of this cell because it's inherited from the previous cell.
                builder.InsertCell();
                builder.Write("Create DateTime");

                builder.EndRow();

                // Set features for the other rows and cells.
                builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
                builder.CellFormat.Width = 100.0;
                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

                // Reset height and define a different height rule for table body.
                builder.RowFormat.Height = 15.0;
                builder.RowFormat.HeightRule = HeightRule.Auto;

                // Reset font formatting.
                builder.Font.Size = 10;
                builder.Font.Bold = false;

                foreach (GridViewRow row in UmbracoMembersGridView.Rows)
                {
                    if (row.RowType == DataControlRowType.DataRow)
                    {
                        CheckBox chkRow = (row.Cells[0].FindControl("SelectedCheckBox") as CheckBox);
                        if (chkRow.Checked)
                        {
                            // Build the other cells.
                            builder.InsertCell();
                            builder.Write(row.Cells[1].Text.ToString());

                            builder.InsertCell();
                            builder.Write(row.Cells[2].Text.ToString());

                            builder.InsertCell();
                            builder.CellFormat.Width = 200.0;
                            builder.Write(row.Cells[3].Text.ToString());

                            builder.InsertCell();
                            builder.Write(row.Cells[4].Text.ToString());
                            builder.EndRow();
                        }
                    }
                }

                // Saves the document to the local file system.
                string fname = System.Guid.NewGuid().ToString() + "." + GetSaveFormat(ExportTypeDropDown.SelectedValue);
                doc.Save(Server.MapPath("~/App_Data/") + fname);
                Response.Clear();
                Response.Buffer = true;
                Response.AddHeader("content-disposition",
                    $"attachment;filename=ExportedFile_{DateTime.Now.Day}_{DateTime.Now.Month}_{DateTime.Now.Year}_{DateTime.Now.Hour}_{DateTime.Now.Minute}_{DateTime.Now.Second}_{DateTime.Now.Millisecond}.{GetSaveFormat(ExportTypeDropDown.SelectedValue)}");
                Response.Charset = "";
                Response.ContentType = "application/pdf";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.ContentType = "Application/" + GetSaveFormat(ExportTypeDropDown.SelectedValue);

                // Get the physical path to the file.
                string FilePath = MapPath("~/App_Data/" + fname);

                // Write the file directly to the HTTP content output stream.
                Response.WriteFile(FilePath);
                Response.Flush();

                // Delete the file as its already in stream and available for user to download/save/view.
                FileInfo file = new FileInfo(FilePath);
                if (file.Exists)
                {
                    file.Delete();
                }
            }
            catch (Exception ex)
            {
                ErrorLabel.Text = ex.ToString();
                ErrorLabel.Visible = true;
            }
        }

        // Get save formats by their respective file extensions. 
        private string GetSaveFormat(string format)
        {
            try
            {
                string saveOption = SaveFormat.Pdf.ToString();
                switch (format)
                {
                    case "Pdf":
                        saveOption = SaveFormat.Pdf.ToString(); break;
                    case "Doc":
                        saveOption = SaveFormat.Doc.ToString(); break;
                    case "Docx":
                        saveOption = SaveFormat.Docx.ToString(); break;
                    case "Odt":
                        saveOption = SaveFormat.Odt.ToString(); break;
                    case "Xps":
                        saveOption = SaveFormat.Xps.ToString(); break;
                    case "Tiff":
                        saveOption = SaveFormat.Tiff.ToString(); break;
                    case "Png":
                        saveOption = SaveFormat.Png.ToString(); break;
                    case "Jpeg":
                        saveOption = SaveFormat.Jpeg.ToString(); break;
                    // The "SaveFormat" property contains more supported save formats.
                }

                return saveOption;
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

        protected void Page_PreRender(object sender, EventArgs e)
        {
            if (ErrorLabel.Visible)
            {
                ErrorLabel.Text = "<br>" + ErrorLabel.Text + "<br>";
            }
        }
    }
}