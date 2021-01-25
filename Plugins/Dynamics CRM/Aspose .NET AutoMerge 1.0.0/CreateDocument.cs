using Aspose.Words;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.AutoMerge
{
    public class CreateDocument : CodeActivity
    {
        [Input("Enable Logging")]
        [Default("False")]
        public InArgument<bool> EnableLogging { get; set; }

        [Input("Log File Directory")]
        [Default("C:\\Aspose Logs")]
        public InArgument<string> LogFile { get; set; }

        [Input("Document Template")]
        [ReferenceTarget("aspose_documenttemplate")]
        public InArgument<EntityReference> DocumentTemplateId { get; set; }

        [Input("License File Path (Optional)")]
        public InArgument<string> LicenseFile { get; set; }

        [Output("Attachment")]
        [ReferenceTarget("annotation")]
        public OutArgument<EntityReference> OutputAttachmentId { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            EntityReference Template = DocumentTemplateId.Get(executionContext);
            Boolean Logging = EnableLogging.Get(executionContext);
            string LicenseFilePath = LicenseFile.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
            OutputAttachmentId.Set(executionContext, new EntityReference("annotation", Guid.Empty));
            try
            {
                if (Logging)
                    Log("Workflow Executed", LogFilePath);
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                string PrimaryEntityName = context.PrimaryEntityName;
                Guid PrimaryEntityId = context.PrimaryEntityId;
                try
                {
                    if (Logging)
                        Log("Enable Licensing", LogFilePath);
                    if (LicenseFilePath != "" && File.Exists(LicenseFilePath))
                    {
                        License Lic = new License();
                        Lic.SetLicense(LicenseFilePath);
                        if (Logging)
                            Log("License Set", LogFilePath);
                    }
                }
                catch (Exception ex)
                {
                    Log("Error while applying license: " + ex.Message, LogFilePath);
                }
                QueryExpression RetrieveNoteQuery = new QueryExpression("annotation");
                RetrieveNoteQuery.ColumnSet = new ColumnSet(new string[] { "subject", "documentbody" });
                RetrieveNoteQuery.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, Template.Id));
                if (Logging)
                    Log("Executing Query to retrieve Template Attachment", LogFilePath);
                EntityCollection Notes = service.RetrieveMultiple(RetrieveNoteQuery);
                if (Logging)
                    Log("Attachment Retrieved Successfully", LogFilePath);
                if (Notes.Entities.Count > 0)
                {
                    Entity Note = Notes[0];
                    string FileName = "";
                    if (Note.Contains("subject"))
                        FileName = Note["subject"].ToString();
                    if (Note.Contains("documentbody"))
                    {
                        if (Logging)
                            Log("Attachment Read Successfully", LogFilePath);
                        byte[] DocumentBody = Convert.FromBase64String(Note["documentbody"].ToString());
                        MemoryStream fileStream = new MemoryStream(DocumentBody);
                        if (Logging)
                            Log("Reading Document in Aspose.Words", LogFilePath);
                        Document doc = new Document(fileStream);
                        if (Logging)
                            Log("Getting Fields list", LogFilePath);

                        string[] fields = doc.MailMerge.GetFieldNames();
                        if (Logging)
                            Log("Getting list of fields for entity", LogFilePath);
                        Entity PrimaryEntity = service.Retrieve(PrimaryEntityName, PrimaryEntityId, new ColumnSet(fields));
                        if (Logging)
                            Log("Retrieved Contact entity", LogFilePath);
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
                            if (Logging)
                                Log("Executing Mail Merge", LogFilePath);
                            doc.MailMerge.Execute(fields, values);
                            MemoryStream UpdateDoc = new MemoryStream();
                            if (Logging)
                                Log("Saving Document", LogFilePath);
                            doc.Save(UpdateDoc, SaveFormat.Docx);
                            byte[] byteData = UpdateDoc.ToArray();

                            // Encode the data using base64.
                            string encodedData = System.Convert.ToBase64String(byteData);

                            if (Logging)
                                Log("Creating Attachment", LogFilePath);
                            Entity NewNote = new Entity("annotation");

                            // Im going to add Note to entity.
                            NewNote.Attributes.Add("objectid", new EntityReference(PrimaryEntityName, PrimaryEntityId));
                            NewNote.Attributes.Add("subject", FileName);

                            // Set EncodedData to Document Body.
                            NewNote.Attributes.Add("documentbody", encodedData);

                            // Set the type of attachment.
                            NewNote.Attributes.Add("mimetype", @"application/vnd.openxmlformats-officedocument.wordprocessingml.document");
                            NewNote.Attributes.Add("notetext", "Document Created using template");

                            // Set the File Name.
                            NewNote.Attributes.Add("filename", FileName);
                            Guid NewNoteId = service.Create(NewNote);
                            OutputAttachmentId.Set(executionContext, new EntityReference("annotation", NewNoteId));
                            if (Logging)
                                Log("Successfull", LogFilePath);
                        }
                    }
                }

                if (Logging)
                    Log("Workflow Executed Successfully", LogFilePath);
            }
            catch (Exception ex)
            {
                Log(ex.Message, LogFilePath);
            }
        }

        private void Log(string Message, string LogFilePath)
        {
            if (LogFilePath == "")
                File.AppendAllText("C:\\Aspose Logs\\Aspose.AutoMerge.CreateDocument.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            else
                File.AppendAllText(LogFilePath + "\\Aspose.AutoMerge.CreateDocument.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
        }
    }
}
