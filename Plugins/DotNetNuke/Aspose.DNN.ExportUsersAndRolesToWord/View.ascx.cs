/*
' Copyright (c) 2015  Aspose.com
'  All rights reserved.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
' TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
' THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
' CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
' 
*/

using System;
using DotNetNuke.Security;
using DotNetNuke.Services.Exceptions;
using DotNetNuke.Entities.Modules;
using DotNetNuke.Entities.Modules.Actions;
using DotNetNuke.Services.Localization;
using Aspose.Words;
using System.IO;
using DotNetNuke.Entities.Users;
using System.Collections.Generic;
using System.Web.UI.WebControls;
using System.Collections;
using System.Data;
using System.Text;


namespace Aspose.DNN.ExportUsersAndRolesToWord
{
    /// -----------------------------------------------------------------------------
    /// <summary>
    /// The View class displays the content
    /// 
    /// Typically your view control would be used to display content or functionality in your module.
    /// 
    /// View may be the only control you have in your project depending on the complexity of your module
    /// 
    /// Because the control inherits from Aspose.DNN.ExportUsersAndRolesToWordModuleBase you have access to any custom properties
    /// defined there, as well as properties from DNN such as PortalId, ModuleId, TabId, UserId and many more.
    /// 
    /// </summary>
    /// -----------------------------------------------------------------------------    
    
    public partial class View : ExportUsersAndRolesToWordModuleBase, IActionable
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                ArrayList dnnUsersArrayList = UserController.GetUsers(PortalId);
                if (dnnUsersArrayList.Count == 0)
                {
                    ExportButton.Visible = false;
                    ExportTypeDropDown.Visible = false;
                }

                ArrayList stuffedUsers = new ArrayList();

                DataTable output = new DataTable("ASA");
                output.Columns.Add("DisplayName");
                output.Columns.Add("Email");
                output.Columns.Add("Roles");
                output.Columns.Add("UserID");


                foreach (UserInfo user in dnnUsersArrayList)
                {

                    string roles = string.Join(",", user.Roles);

                    DataRow dr;
                    dr = output.NewRow();
                    dr["DisplayName"] = user.DisplayName;
                    dr["Email"] = user.Email;
                    dr["Roles"] = roles;
                    dr["UserID"] = user.UserID;

                    output.Rows.Add(dr);


                }

                UsersGridView.DataSource = output;
                if (!IsPostBack)
                    UsersGridView.DataBind();
            }
            catch (Exception exc) //Module failed to load
            {
                Exceptions.ProcessModuleLoadException(this, exc);
            }
        }

        public ModuleActionCollection ModuleActions
        {
            get
            {
                var actions = new ModuleActionCollection
                    {
                        {
                            GetNextActionID(), Localization.GetString("EditModule", LocalResourceFile), "", "", "",
                            EditUrl(), false, SecurityAccessLevel.Edit, true, false
                        }
                    };
                return actions;
            }
        }

        protected void ExportButton_Click(object sender, EventArgs e)
        {

            string format = ExportTypeDropDown.SelectedValue;
            List<UserInfo> usersList = new List<UserInfo>();
            //CType(GridView1.Rows(i).FindControl("cbMove"), CheckBox).Checked()

            foreach (GridViewRow row in UsersGridView.Rows)
            {
                if (row.RowType == DataControlRowType.DataRow)
                {
                    CheckBox chkRow = (row.Cells[0].FindControl("SelectedCheckBox") as CheckBox);
                    if (chkRow.Checked)
                    {
                        int userId = Convert.ToInt32(UsersGridView.DataKeys[row.RowIndex].Value.ToString());
                        usersList.Add(UserController.GetUserById(PortalId, userId));
                    }
                }
            }

            if(usersList.Count == 0)
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

                string content_header_rich = "<tr><th>Name</th><th>Email</th><th>Roles</th></tr>";
                string content_rows_rich = "";
                

                foreach (UserInfo user in usersList)
                {
                    content_rows_rich = string.Concat(content_rows_rich,"<tr><td>" + user.DisplayName + "</td><td>" + user.Email + "</td><td>" + string.Join(",", user.Roles) + "</td></tr>");                    
                }

                string content_rich = "<html><h2>Exported Users and Roles</h2><table width=\"100%\" style='border: 1px solid black; border-collapse: collapse;' >" + content_header_rich + content_rows_rich + "</table></html>";

                if(format != "txt")
                {
                    MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(content_rich));
                    Document doc = new Document(stream);
                    string fileName = GetOutputFileName(format);                    
                    doc.Save(Response, fileName, ContentDisposition.Inline, null);
                    Response.End();
                }
                else
                {
                    Document doc = new Document();
                    DocumentBuilder builder = new DocumentBuilder(doc);

                    string delim = "\t";

                    builder.Writeln("Name" + delim + "Email" + delim + "Roles");

                    foreach (UserInfo user in usersList)
                    {
                        string row = user.DisplayName + delim + user.Email + delim + string.Join(",", user.Roles);
                        builder.Writeln(row);
                    }


                    string filename = GetOutputFileName(format);
                    doc.Save(MapPath("./AsposeOutput/" + filename));

                    DownloadNow(filename);

                }                
            }

            


        }

        private string GetOutputFileName(string extension)
        {
            string name = System.Guid.NewGuid().ToString() + "." + extension;
            return name;
        }

        protected void DownloadNow(string file)
        {                        
            string fileName = "./AsposeOutput/" + file;

            FileStream fs = new FileStream(MapPath(fileName), FileMode.Open);
            long cntBytes = new FileInfo(MapPath(fileName)).Length;
            byte[] byteArray = new byte[Convert.ToInt32(cntBytes)];
            fs.Read(byteArray, 0, Convert.ToInt32(cntBytes));
            fs.Close();


            

            if (byteArray != null)
            {
                this.Response.Clear();
                this.Response.ContentType = "text/plain";
                this.Response.AddHeader("content-disposition", "attachment;filename=" + file);
                this.Response.BinaryWrite(byteArray);
                this.Response.End();
                this.Response.Flush();
                this.Response.Close();
            }
            File.Delete(Server.MapPath(fileName));
        }
    }
}