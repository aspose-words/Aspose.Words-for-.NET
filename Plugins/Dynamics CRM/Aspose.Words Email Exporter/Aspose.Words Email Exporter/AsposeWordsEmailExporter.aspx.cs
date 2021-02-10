using Aspose.Words;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.ServiceModel.Description;

namespace Aspose.Words_Email_Exporter
{
    public partial class AsposeWordsEmailExporter : System.Web.UI.Page
    {
        OrganizationServiceProxy Service;
        Guid EntityId = Guid.Empty;
        string EntityName;
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                LBL_Message.Text = "";
                EntityId = Request.QueryString["id"] == null ? Guid.Empty : new Guid(Request.QueryString["id"]);
                string OrgName = Request.QueryString["orgname"] == null ? "" : Request.QueryString["orgname"];
                EntityName = Request.QueryString["typename"] == null ? "" : Request.QueryString["typename"];
                if (EntityId != Guid.Empty && EntityName != "" && OrgName != "")
                {
                    CreateService(ConfigurationManager.AppSettings["ServerName"], OrgName, ConfigurationManager.AppSettings["Login"],
                         ConfigurationManager.AppSettings["Password"], ConfigurationManager.AppSettings["Domain"]);
                    try
                    {
                        string LicenseFilePath = ConfigurationManager.AppSettings["LicenseFilePath"];

                        // Check for license and apply if exists.
                        if (File.Exists(LicenseFilePath))
                        {
                            License license = new License();
                            license.SetLicense(LicenseFilePath);
                        }
                    }
                    catch { }
                    if (!IsPostBack)
                    {
                        Entity Email = Service.Retrieve("email", EntityId, new ColumnSet(new string[] { "subject" }));
                        TXT_FileName.Text = Email.Contains("subject") ? Email["subject"].ToString() : "";
                    }
                }
                else
                {
                    LBL_Message.Text = "Parameters are not correct, Please save the record first.";
                }
            }
            catch (Exception ex)
            {
                LBL_Message.Text = "Error has occured. Details: " + ex.Message;
            }
        }

        protected void BTN_Generate_Click(object sender, EventArgs e)
        {
            try
            {
                if (ValidateFields())
                {
                    string Action = DDL_Action.SelectedValue;
                    string Format = DDL_FileFormat.SelectedValue.ToLower();
                    string FileName = TXT_FileName.Text;

                    Entity Email = Service.Retrieve("email", EntityId, new ColumnSet(true));
                    string EmailBody = Email.Contains("description") ? Email["description"].ToString() : "";

                    string UpdatedEmailBody = ReplaceImagesInBody(EmailBody);


                    DocumentBuilder MyDoc = new DocumentBuilder();
                    MyDoc.InsertHtml(UpdatedEmailBody);
                    Document doc = MyDoc.Document;
                    MemoryStream UpdateDoc = new MemoryStream();

                    switch (Format.ToLower())
                    {
                        case "bmp":
                            doc.Save(UpdateDoc, SaveFormat.Bmp);
                            break;
                        case "doc":
                            doc.Save(UpdateDoc, SaveFormat.Doc);
                            break;
                        case "docx":
                            doc.Save(UpdateDoc, SaveFormat.Docx);
                            break;
                        case "html":
                            doc.Save(UpdateDoc, SaveFormat.Html);
                            break;
                        case "jpeg":
                            doc.Save(UpdateDoc, SaveFormat.Jpeg);
                            break;
                        case "pdf":
                            doc.Save(UpdateDoc, SaveFormat.Pdf);
                            break;
                        case "png":
                            doc.Save(UpdateDoc, SaveFormat.Png);
                            break;
                        case "rtf":
                            doc.Save(UpdateDoc, SaveFormat.Rtf);
                            break;
                        case "text":
                        case "txt":
                            doc.Save(UpdateDoc, SaveFormat.Text);
                            break;
                        default:
                            doc.Save(UpdateDoc, SaveFormat.Docx);
                            break;
                    }

                    if (Action == "Download")
                    {
                        Response.Clear();
                        Response.ContentType = "Application/msword";
                        Response.AddHeader("Content-Disposition", "attachment; filename=" + FileName + "." + Format);
                        Response.BinaryWrite(UpdateDoc.ToArray());
                        // myMemoryStream.WriteTo(Response.OutputStream); //works too
                        Response.Flush();
                        Response.Close();
                        Response.End();
                    }
                    if (Action == "Attach to This Email")
                    {
                        if (((OptionSetValue)Email["statecode"]).Value == 1)
                        {
                            LBL_Message.Text = "Email is closed, Download the File instead";
                            Response.Clear();
                            Response.ContentType = "Application/msword";
                            Response.AddHeader("Content-Disposition", "attachment; filename=" + FileName + "." + Format);
                            Response.BinaryWrite(UpdateDoc.ToArray());
                            // myMemoryStream.WriteTo(Response.OutputStream); //works too
                            Response.Flush();
                            Response.Close();
                            Response.End();
                        }
                        else
                        {
                            Entity NewNote = new Entity("activitymimeattachment");
                            byte[] byteData = UpdateDoc.ToArray();

                            // Encode the data using base64.
                            string encodedData = System.Convert.ToBase64String(byteData);

                            // Add a Note to the entity.
                            NewNote.Attributes.Add("objectid", new EntityReference("email", EntityId));
                            NewNote.Attributes.Add("objecttypecode", "email");
                            NewNote.Attributes.Add("subject", FileName + "." + Format);

                            // Set EncodedData to Document Body.
                            NewNote.Attributes.Add("body", encodedData);

                            // Set the type of attachment.
                            NewNote.Attributes.Add("mimetype", @"application/vnd.openxmlformats-officedocument.wordprocessingml.document");
                            //NewNote.Attributes.Add("notetext", FileName);

                            // Set the filename.
                            NewNote.Attributes.Add("filename", FileName + "." + Format);
                            Guid NewNoteId = Service.Create(NewNote);
                        }
                    }
                }
                else
                {
                    LBL_Message.Text = "Please enter the fields";
                }
            }
            catch (Exception ex)
            {
                LBL_Message.Text = "Error has occured. Details: " + ex.Message;
            }
        }

        private string ReplaceImagesInBody(string EmailBody)
        {
            try
            {
                List<string> Images = new List<string>();
                int startIndex = EmailBody.IndexOf("<img ");
                if (startIndex < 0)
                    return EmailBody;
                int endIndex = EmailBody.IndexOf(">", startIndex);
                while (startIndex >= 0 && endIndex >= 0)
                {
                    string ImageTag = EmailBody.Substring(startIndex, endIndex - startIndex + 1);
                    Images.Add(ImageTag);
                    startIndex = EmailBody.IndexOf("<img ", endIndex);
                    if (startIndex >= 0)
                        endIndex = EmailBody.IndexOf(">", startIndex);
                }
                foreach (string Image in Images)
                {
                    try
                    {
                        int start = 0;
                        int end = 0;
                        if (Image.ToLower().Contains("src="))
                        {
                            if (Image.ToLower().Contains("src='"))
                            {
                                start = Image.IndexOf("src='");
                                end = Image.IndexOf("'", start + 5);
                            }
                            else
                            {
                                start = Image.IndexOf("src=\"");
                                end = Image.IndexOf("\"", start + 5);
                            }
                            string src = Image.Substring(start, end - start + 1);
                            if (src.Contains("attachmentid="))
                            {
                                string attachmentid = src.Substring(src.ToLower().IndexOf("attachmentid="));
                                attachmentid = attachmentid.ToLower().Replace("attachmentid=", "").Replace("'", "").Replace("\"", "");
                                Guid AttachmentId = new Guid(attachmentid);
                                Entity Attachment = Service.Retrieve("activitymimeattachment", AttachmentId, new ColumnSet(true));
                                string ImageBytes = Attachment.Contains("body") ? Attachment["body"].ToString() : "";
                                string Mimetype = Attachment.Contains("mimetype") ? Attachment["mimetype"].ToString() : "";
                                string FileName = Attachment.Contains("filename") ? Attachment["filename"].ToString() : "";

                                string UpdatedSrc = "src=\"data:" + Mimetype + ";base64," + ImageBytes + "\"";
                                string UpdatedImage = Image.Replace(src, UpdatedSrc);
                                EmailBody = EmailBody.Replace(Image, UpdatedImage);
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                LBL_Message.Text = "Error replacing images in email body. Details: " + ex.Message;
            }
            return EmailBody;
        }

        private bool ValidateFields()
        {
            if (String.IsNullOrEmpty(DDL_FileFormat.SelectedValue))
                return false;
            if (String.IsNullOrEmpty(DDL_Action.SelectedValue))
                return false;
            if (String.IsNullOrEmpty(TXT_FileName.Text))
                return false;

            return true;
        }

        private OrganizationServiceProxy CreateService(string serverName, string OrganizationName, string Login, string Password, string Domain)
        {
            try
            {
                serverName = serverName.Contains("http://") || serverName.Contains("https://") ? serverName.Replace("http://", "").Replace("https://", "") : serverName;
                Uri oUri = new Uri("http://" + serverName + "/" + OrganizationName + "/XRMServices/2011/Organization.svc");
                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = Domain + "\\" + Login;
                clientCredentials.UserName.Password = Password;
                Service = new OrganizationServiceProxy(oUri, null, clientCredentials, null);
                return Service;
            }
            catch (Exception ex)
            {
                Service = null;
                throw ex;
            }
        }
    }
}