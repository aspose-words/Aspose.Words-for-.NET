using Aspose.Words;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.ServiceModel.Description;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Aspose.NET_OneClick_Word_Document
{
    public partial class OneClick_Word_Page : System.Web.UI.Page
    {
        OrganizationServiceProxy Service;
        string TemplateEntityName = "aspose_oneclickwordconfiguration";
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
                        // Check for license and apply if exists
                        if (File.Exists(LicenseFilePath))
                        {
                            License license = new License();
                            license.SetLicense(LicenseFilePath);
                        }
                    }
                    catch { }
                    if (!IsPostBack)
                    {

                        LoadDataInDropDowns(EntityName);
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

        private void LoadDataInDropDowns(string EntityName)
        {
            int EntityType = GetEntityType(EntityName);
            if (EntityType != 0)
            {
                DDL_Templates.Items.Clear();
                QueryExpression Query = new QueryExpression(TemplateEntityName);
                Query.Criteria.AddCondition(new ConditionExpression("aspose_entity", ConditionOperator.Equal, EntityType));
                Query.ColumnSet = new ColumnSet(new string[] { "aspose_name" });
                EntityCollection Templates = Service.RetrieveMultiple(Query);
                foreach (Entity Template in Templates.Entities)
                {
                    DDL_Templates.Items.Add(new ListItem(Template["aspose_name"].ToString(), Template.Id.ToString()));
                }
            }
        }

        private int GetEntityType(string EntityName)
        {
            switch (EntityName.ToLower())
            {
                case "account":
                    return 1;
                case "contact":
                    return 2;
                case "opportunity":
                    return 3;
                case "lead":
                    return 4;
                case "quote":
                    return 1084;
                default:
                    return 0;
            }
        }

        protected void BTN_Generate_Click(object sender, EventArgs e)
        {
            try
            {
                if (ValidateFields())
                {
                    string SelectedTemplateId = DDL_Templates.SelectedValue;
                    string Action = DDL_Action.SelectedValue;
                    string Format = DDL_FileFormat.SelectedValue.ToLower();
                    string FileName = TXT_FileName.Text;
                    Guid TemplateId = new Guid(SelectedTemplateId);
                    QueryExpression NotesQuery = new QueryExpression("annotation");
                    NotesQuery.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, TemplateId));
                    NotesQuery.ColumnSet = new ColumnSet(new string[] { "subject", "documentbody" });
                    EntityCollection Notes = Service.RetrieveMultiple(NotesQuery);
                    if (Notes.Entities.Count > 0)
                    {
                        Entity Note = Notes[0];
                        if (Note.Contains("documentbody"))
                        {
                            byte[] DocumentBody = Convert.FromBase64String(Note["documentbody"].ToString());
                            MemoryStream fileStream = new MemoryStream(DocumentBody);
                            Document doc = new Document(fileStream);
                            string[] fields = doc.MailMerge.GetFieldNames();
                            Entity PrimaryEntity = Service.Retrieve(EntityName, EntityId, new ColumnSet(fields));
                            if (PrimaryEntity != null)
                            {
                                string[] values = new string[fields.Length];
                                for (int i = 0; i < fields.Length; i++)
                                {
                                    if (PrimaryEntity.Contains(fields[i]))
                                    {
                                        if (PrimaryEntity[fields[i]].GetType() == typeof(OptionSetValue))
                                            values[i] = PrimaryEntity.FormattedValues[fields[i]].ToString();
                                        else if (PrimaryEntity[fields[i]].GetType() == typeof(EntityReference))
                                            values[i] = ((EntityReference)PrimaryEntity[fields[i]]).Name;
                                        else
                                            values[i] = PrimaryEntity[fields[i]].ToString();
                                    }
                                    else
                                        values[i] = "";
                                }
                                doc.MailMerge.Execute(fields, values);
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
                                if (Action == "Attach to Note")
                                {
                                    byte[] byteData = UpdateDoc.ToArray();
                                    // Encode the data using base64.
                                    string encodedData = System.Convert.ToBase64String(byteData);
                                    Entity NewNote = new Entity("annotation");
                                    // Im going to add Note to entity
                                    NewNote.Attributes.Add("objectid", new EntityReference(EntityName, EntityId));
                                    NewNote.Attributes.Add("subject", FileName + "." + Format);

                                    // Set EncodedData to Document Body
                                    NewNote.Attributes.Add("documentbody", encodedData);

                                    // Set the type of attachment
                                    NewNote.Attributes.Add("mimetype", @"application/vnd.openxmlformats-officedocument.wordprocessingml.document");
                                    NewNote.Attributes.Add("notetext", FileName);

                                    // Set the File Name
                                    NewNote.Attributes.Add("filename", FileName + "." + Format);
                                    Guid NewNoteId = Service.Create(NewNote);
                                }
                            }
                        }
                    }
                    else
                    {
                        LBL_Message.Text = "No Attachment found in the selected Configuration";
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

        private bool ValidateFields()
        {
            if (String.IsNullOrEmpty(DDL_Templates.SelectedValue))
                return false;
            if (String.IsNullOrEmpty(DDL_Action.SelectedValue))
                return false;
            if (String.IsNullOrEmpty(DDL_FileFormat.SelectedValue))
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