using Aspose.Words;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.DocumentGenerator.CreateDocumentFromTemplate
{
    public class CreateDocument : CodeActivity
    {
        [Input("Document Template")]
        [ReferenceTarget("aspose_documenttemplate")]
        public InArgument<EntityReference> DocumentTemplateId { get; set; }

        [Input("Contact")]
        [ReferenceTarget("contact")]
        public InArgument<EntityReference> ContactId { get; set; }

        [Input("Enable Logging")]
        public InArgument<bool> EnableLogging { get; set; }

        [Input("License File Path (Optional)")]
        public InArgument<string> LicenseFile { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            try
            {
                EntityReference Template = DocumentTemplateId.Get(executionContext);
                EntityReference Contact = ContactId.Get(executionContext);
                Boolean Logging = EnableLogging.Get(executionContext);
                string LicenseFilePath = LicenseFile.Get(executionContext);

                if (Logging)
                    Log("Workflow Executed");

                // Create a CRM Service in Workflow.
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                QueryExpression RetrieveNoteQuery = new QueryExpression("annotation");
                RetrieveNoteQuery.ColumnSet = new ColumnSet(new string[] { "subject", "documentbody" });
                RetrieveNoteQuery.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, Template.Id));

                if (Logging)
                    Log("Executing Query to retrieve Template Attachment");

                EntityCollection Notes = service.RetrieveMultiple(RetrieveNoteQuery);

                if (Logging)
                    Log("Attachment Retrieved Successfully");

                if (Notes.Entities.Count > 0)
                {
                    Entity Note = Notes[0];
                    string FileName = "";
                    if (Note.Contains("subject"))
                        FileName = Note["subject"].ToString();
                    if (Note.Contains("documentbody"))
                    {
                        if (Logging)
                            Log("Attachment Read Successfully");
                        byte[] DocumentBody = Convert.FromBase64String(Note["documentbody"].ToString());
                        MemoryStream fileStream = new MemoryStream(DocumentBody);
                        try
                        {
                            if (Logging)
                                Log("Enable Licensing");
                            if (LicenseFilePath != "" && File.Exists(LicenseFilePath))
                            {
                                Aspose.Words.License Lic = new License();
                                Lic.SetLicense(LicenseFilePath);
                                if (Logging)
                                    Log("License Set");
                            }
                        }
                        catch (Exception ex)
                        {
                            Log("Error while applying license: " + ex.Message);
                        }

                        if (Logging)
                            Log("Reading Document in Aspose.Words");

                        Document doc = new Document(fileStream);
                        string[] fields = doc.MailMerge.GetFieldNames();

                        if (Logging)
                            Log("Getting list of fields");

                        Entity contact = service.Retrieve("contact", Contact.Id, new ColumnSet(fields));

                        if (Logging)
                            Log("Retrieved Contact entity");

                        if (contact != null)
                        {
                            string[] values = new string[fields.Length];
                            for (int i = 0; i < fields.Length; i++)
                            {
                                if (contact.Contains(fields[i]))
                                {
                                    if (contact[fields[i]].GetType() == typeof(OptionSetValue))
                                        values[i] = contact.FormattedValues[fields[i]].ToString();
                                    else if (contact[fields[i]].GetType() == typeof(EntityReference))
                                        values[i] = ((EntityReference)contact[fields[i]]).Name;
                                    else
                                        values[i] = contact[fields[i]].ToString();
                                }
                                else
                                    values[i] = "";
                            }

                            if (Logging)
                                Log("Executing Mail Merge");

                            doc.MailMerge.Execute(fields, values);
                            MemoryStream UpdateDoc = new MemoryStream();

                            if (Logging)
                                Log("Saving Document");

                            doc.Save(UpdateDoc, SaveFormat.Docx);
                            byte[] byteData = UpdateDoc.ToArray();

                            // Encode the data using base64.
                            string encodedData = System.Convert.ToBase64String(byteData);

                            if (Logging)
                                Log("Creating Attachment");

                            Entity NewNote = new Entity("annotation");
                            // Im going to add Note to entity.
                            NewNote.Attributes.Add("objectid", new EntityReference("contact", Contact.Id));
                            NewNote.Attributes.Add("subject", FileName);

                            // Set EncodedData to Document Body.
                            NewNote.Attributes.Add("documentbody", encodedData);

                            // Set the type of attachment.
                            NewNote.Attributes.Add("mimetype", @"application\ms-word");
                            NewNote.Attributes.Add("notetext", "Document Created using template");

                            // Set the File Name.
                            NewNote.Attributes.Add("filename", FileName);
                            service.Create(NewNote);

                            if (Logging)
                                Log("Successfull");
                        }
                    }
                }
                if (Logging)
                    Log("Workflow Executed Successfully");
            }
            catch (Exception ex)
            {
                Log(ex.Message);
            }
        }

        private void Log(string Message)
        {
            File.AppendAllText("C:\\Aspose Logs\\Aspose.DocumentGenerator.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
        }
    }
}
