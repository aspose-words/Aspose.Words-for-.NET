using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Telerik.Sitefinity.Security;
using Telerik.Sitefinity.Security.Model;
using Aspose.Words;
using Aspose.Words.Saving;
using System.IO;
using System.Text;

namespace Aspose.Sitefinity.ExportUsersToWord
{
    public partial class AsposeExportUsersToWord : System.Web.UI.UserControl
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            UserManager userManager = UserManager.GetManager();
            List<User> users = userManager.GetUsers().ToList();

            DataTable output = new DataTable("ASA");
            output.Columns.Add("LastName");
            output.Columns.Add("FirstName");
            output.Columns.Add("Email");
            output.Columns.Add("Username");
            output.Columns.Add("UserID");


            foreach (User user in users)
            {
                UserProfileManager profileManager = UserProfileManager.GetManager();
                SitefinityProfile profile = profileManager.GetUserProfile<SitefinityProfile>(user);

                DataRow dr;
                dr = output.NewRow();
                dr["FirstName"] = profile.FirstName;
                dr["Username"] = user.UserName;
                dr["Email"] = user.Email;
                dr["LastName"] = profile.LastName;
                dr["UserID"] = user.Id;

                output.Rows.Add(dr);


            }

            Document d = new Document();
            

            


            SitefinityUsersGridView.DataSource = output;
            //if (!IsPostBack)
            SitefinityUsersGridView.DataBind();
        }

        protected void ExportButton_Click(object sender, EventArgs e)
        {
            string format = ExportTypeDropDown.SelectedValue;

            DataTable selectedUsers = new DataTable("SU");
            selectedUsers.Columns.Add("UserID");
            selectedUsers.Columns.Add("DisplayName");
            selectedUsers.Columns.Add("Email");
            selectedUsers.Columns.Add("Username");

            foreach (GridViewRow row in SitefinityUsersGridView.Rows)
            {
                if (row.RowType == DataControlRowType.DataRow)
                {
                    CheckBox chkRow = (row.Cells[0].FindControl("SelectedCheckBox") as CheckBox);
                    if (chkRow.Checked)
                    {
                        string email = SitefinityUsersGridView.DataKeys[row.RowIndex].Value.ToString();
                        var userMan = UserManager.GetManager();
                        User user = userMan.GetUserByEmail(email);

                        UserProfileManager profileManager = UserProfileManager.GetManager();
                        SitefinityProfile profile = profileManager.GetUserProfile<SitefinityProfile>(user);

                        DataRow dr;
                        dr = selectedUsers.NewRow();
                        dr["UserID"] = user.Id;
                        dr["DisplayName"] = profile.FirstName + " " + profile.LastName;
                        dr["Email"] = user.Email;
                        dr["Username"] = user.UserName;
                        selectedUsers.Rows.Add(dr);
                    }
                }
            }

            if (selectedUsers.Rows.Count == 0)
            {
                NoRowSelectedErrorDiv.Visible = true;
            }
            else
            {
                // Check for license and apply if exists
                string licenseFile = Server.MapPath("~/App_Data/Aspose.Words.lic");
                if (File.Exists(licenseFile))
                {
                    License license = new License();
                    license.SetLicense(licenseFile);
                }

                string content_header_rich = "<tr><th>UserID</th><th>Display Name</th><th>Email</th><th>Username</th></tr>";
                string content_rows_rich = "";


                foreach (DataRow user in selectedUsers.Rows)
                {
                    content_rows_rich = string.Concat(content_rows_rich, "<tr><td>" + user["UserID"].ToString() + "</td><td>" + user["DisplayName"].ToString() + "</td><td>" + user["Email"].ToString() + "</td><td>" + user["Username"].ToString() + "</td></tr>");
                }

                string content_rich = "<html><h2>Exported Users</h2><table width=\"100%\" style='border: 1px solid black; border-collapse: collapse;' >" + content_header_rich + content_rows_rich + "</table></html>";
                string fileName = System.Guid.NewGuid().ToString() + "." + format;

                if (format != "txt")
                {
                    MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(content_rich));
                    Document doc = new Document(stream);                    
                    doc.Save(Response, fileName, ContentDisposition.Inline, null);
                    Response.End();
                }
                else
                {
                    Document doc = new Document();
                    DocumentBuilder builder = new DocumentBuilder(doc);

                    string delim = "\t";

                    builder.Writeln("UserID" + delim + "DisplayName" + delim + "Email" + delim + "Username");

                    foreach (DataRow user in selectedUsers.Rows)
                    {
                        string row = user["UserID"].ToString() + delim + user["DisplayName"].ToString() + delim + user["Email"].ToString() + delim + user["Username"].ToString();
                        builder.Writeln(row);
                    }

                    
                    doc.Save(this.Response, fileName, ContentDisposition.Attachment, GetSaveFormat(format));
                    Response.End();
                }                
            }
        }

        

        private SaveOptions GetSaveFormat(string format)
        {
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);


            switch (format)
            {
                case "pdf":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf); break;
                case "doc":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Doc); break;
                case "docx":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx); break;
                case "dot":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Dot); break;
                case "dotx":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Dotx); break;
                case "docm":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docm); break;
                case "dotm":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Dotm); break;
                case "odt":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Odt); break;
                case "ott":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Ott); break;
                case "rtf":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Rtf); break;
                case "txt":
                    saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Text); break;
            }

            return saveOptions;
        }
    }
}