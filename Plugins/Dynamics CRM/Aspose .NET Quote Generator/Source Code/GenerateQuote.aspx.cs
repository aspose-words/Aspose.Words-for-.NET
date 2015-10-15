using Aspose.Words;
using Aspose.Words.Saving;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.ServiceModel.Description;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Aspose.QuoteGenerator
{
    public partial class GenerateQuote : System.Web.UI.Page
    {
        OrganizationServiceProxy Service;
        string QuoteTemplateEntityName = "aspose_quotetemplate";
        Guid QuoteId;
        string QuoteName;
        protected void Page_Load(object sender, EventArgs e)
        {
            QuoteId = Request.QueryString["id"] == null ? Guid.Empty : new Guid(Request.QueryString["id"]);
            string OrgName = Request.QueryString["orgname"] == null ? "" : Request.QueryString["orgname"];
            CreateService(ConfigurationManager.AppSettings["ServerName"], OrgName, ConfigurationManager.AppSettings["Login"],
                ConfigurationManager.AppSettings["Password"], ConfigurationManager.AppSettings["Domain"]);
            if (!IsPostBack)
                LoadTemplates();
        }

        private void LoadTemplates()
        {
            WhoAmIResponse res = (WhoAmIResponse)Service.Execute(new WhoAmIRequest());
            QueryExpression query = new QueryExpression(QuoteTemplateEntityName);
            query.ColumnSet = new ColumnSet(true);
            EntityCollection Templates = Service.RetrieveMultiple(query);
            DDL_Templates.Items.Clear();
            foreach (Entity Template in Templates.Entities)
            {
                if (Template.Contains("aspose_name"))
                    DDL_Templates.Items.Add(new ListItem(Template["aspose_name"].ToString(), Template.Id.ToString()));
            }
            LoadTemplateBody();
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

        protected void DDL_Templates_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadTemplateBody();
        }

        private void LoadTemplateBody()
        {
            if (DDL_Templates.SelectedIndex != -1 && DDL_Templates.Items.Count > 0)
            {
                Guid SelectedTemplateId = new Guid(DDL_Templates.SelectedValue);
                Entity QuoteTemplate = Service.Retrieve(QuoteTemplateEntityName, SelectedTemplateId, new ColumnSet(new string[] { "aspose_body" }));
                if (QuoteTemplate != null && QuoteTemplate.Contains("aspose_body"))
                {
                    string body = QuoteTemplate["aspose_body"].ToString();
                    ReplaceTags(body);

                }
            }
        }

        private void ReplaceTags(string body)
        {
            string TagStart = "&lt;MERGEFIELD&gt;";
            string TagEnd = "&lt;/MERGEFIELD&gt;";
            int index1 = 0;
            int index2 = 0;
            List<string> Fields = new List<string>();
            List<string> FieldsLower = new List<string>();
            while (index1 < body.Length)
            {
                index1 = body.IndexOf(TagStart, index2);
                index2 = body.IndexOf(TagEnd, index2 + 1);
                if (index1 == -1 || index2 == -1)
                    break;
                var FieldName = body.Substring(index1 + 18, index2 - index1 - 18);
                Fields.Add(FieldName);
            }
            if (Fields.Count > 0)
            {
                List<string> FieldsNameLower = new List<string>();
                foreach (string FieldLower in Fields)
                    FieldsNameLower.Add(FieldLower.ToLower());
                FieldsNameLower.Add("name");
                Entity Quote = Service.Retrieve("quote", QuoteId, new ColumnSet(FieldsNameLower.ToArray()));
                foreach (string Field in Fields)
                {
                    string Value = "";
                    if (Quote[Field.ToLower()].GetType() == typeof(OptionSetValue))
                        Value = Quote.FormattedValues[Field.ToLower()].ToString();
                    else if (Quote[Field.ToLower()].GetType() == typeof(EntityReference))
                        Value = ((EntityReference)Quote[Field.ToLower()]).Name;
                    else
                        Value = Quote[Field.ToLower()].ToString();

                    if (Quote.Contains(Field.ToLower()))
                        body = body.Replace("&lt;MERGEFIELD&gt;" + Field + "&lt;/MERGEFIELD&gt;", Value);
                }
                QuoteName = Quote["name"].ToString();
            }

            editor1.InnerHtml = body;
        }

        protected void BTN_Download_Click(object sender, EventArgs e)
        {
            Stream stream = GenerateStreamFromString(editor1.InnerText);
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFormat = LoadFormat.Html;

            Document myDoc = new Document(stream, loadOptions);

            MemoryStream memStream = new MemoryStream();
            myDoc.Save(memStream, SaveOptions.CreateSaveOptions(SaveFormat.Docx));

            Response.Clear();
            Response.ContentType = "Application/msword";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + (String.IsNullOrEmpty(QuoteName) ? "Aspose .NET Quote Generator" : QuoteName) + ".docx");
            Response.BinaryWrite(memStream.ToArray());
            // myMemoryStream.WriteTo(Response.OutputStream); //works too
            Response.Flush();
            Response.Close();
            Response.End();

        }
        protected Stream GenerateStreamFromString(string s)
        {
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }
        protected void BTN_Attach_Click(object sender, EventArgs e)
        {
            QuoteName = (String.IsNullOrEmpty(QuoteName) ? "Aspose .NET Quote Generator" : QuoteName);
            Stream stream = GenerateStreamFromString(editor1.InnerText);
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFormat = LoadFormat.Html;

            Document myDoc = new Document(stream, loadOptions);

            MemoryStream memStream = new MemoryStream();
            myDoc.Save(memStream, SaveOptions.CreateSaveOptions(SaveFormat.Docx));


            byte[] byteData = memStream.ToArray();
            // Encode the data using base64.
            string encodedData = System.Convert.ToBase64String(byteData);

            Entity NewNote = new Entity("annotation");
            // Im going to add Note to entity
            NewNote.Attributes.Add("objectid", new EntityReference("quote", QuoteId));
            NewNote.Attributes.Add("subject", QuoteName);

            // Set EncodedData to Document Body
            NewNote.Attributes.Add("documentbody", encodedData);

            // Set the type of attachment
            NewNote.Attributes.Add("mimetype", @"application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            NewNote.Attributes.Add("notetext", "Document Created using template");

            // Set the File Name
            NewNote.Attributes.Add("filename", QuoteName + ".docx");
            Guid NewNoteId = Service.Create(NewNote);
            if (NewNoteId != Guid.Empty)
                LBL_Message.Text = "Successfully added to Quote";
        }
    }
}