using Aspose.Words;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.DuplicateDocument
{
    public class DuplicateDocument : CodeActivity
    {
        [RequiredArgument]
        [Input("Enable Logging")]
        [Default("False")]
        public InArgument<bool> EnableLogging { get; set; }

        [RequiredArgument]
        [Input("Log File Directory")]
        [Default("C:\\Aspose Logs")]
        public InArgument<string> LogFile { get; set; }

        [RequiredArgument]
        [Input("Detect Duplicates in")]
        public InArgument<int> DetectIn { get; set; }

        [Input("License File Path (Optional)")]
        public InArgument<string> LicenseFile { get; set; }

        [Output("Attachment")]
        [ReferenceTarget("annotation")]
        public OutArgument<EntityReference> OutputAttachmentId { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            bool Logging = EnableLogging.Get(executionContext);
            string LicenseFilePath = LicenseFile.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
            int detectIn = DetectIn.Get(executionContext);
            OutputAttachmentId.Set(executionContext, new EntityReference("annotation", Guid.Empty));

            if (Logging)
                Log("Execution Started", LogFilePath);

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

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

            if (detectIn == 0) // under this record
            {
                Guid ThisRecordId = context.PrimaryEntityId;
                string RecordType = context.PrimaryEntityName;
                Document Result = new Document();
                DocumentBuilder ResultWriter = new DocumentBuilder(Result);
                if (Logging)
                    Log("Working under all attachments under this record", LogFilePath);
                QueryExpression RetrieveNoteQuery = new QueryExpression("annotation");
                RetrieveNoteQuery.ColumnSet = new ColumnSet(new string[] { "filename", "subject", "documentbody" });
                RetrieveNoteQuery.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, ThisRecordId));
                if (Logging)
                    Log("Executing Query to retrieve All Notes within this record", LogFilePath);
                EntityCollection Notes = service.RetrieveMultiple(RetrieveNoteQuery);

                foreach (Entity Note in Notes.Entities)
                {
                    try
                    {
                        if (Note.Contains("documentbody"))
                        {
                            string FileName = "";
                            if (Note.Contains("filename"))
                                FileName = Note["filename"].ToString();

                            byte[] DocumentBody = Convert.FromBase64String(Note["documentbody"].ToString());
                            MemoryStream fileStream = new MemoryStream(DocumentBody);
                            Document doc = new Document(fileStream);

                            ResultWriter.Writeln("Comparing Document: " + FileName);
                            ResultWriter.StartTable();

                            foreach (Entity OtherNote in Notes.Entities)
                            {
                                if (OtherNote.Id != Note.Id)
                                {
                                    if (OtherNote.Contains("documentbody"))
                                    {
                                        string OtherFileName = "";
                                        if (OtherNote.Contains("filename"))
                                            OtherFileName = OtherNote["filename"].ToString();
                                        byte[] OtherDocumentBody = Convert.FromBase64String(OtherNote["documentbody"].ToString());
                                        MemoryStream fileStream2 = new MemoryStream(OtherDocumentBody);
                                        Document doc2 = new Document(fileStream);

                                        ResultWriter.InsertCell();
                                        ResultWriter.Write(OtherFileName);

                                        doc.Compare(doc2, "a", DateTime.Now);
                                        if (doc.Revisions.Count == 0)
                                        {
                                            ResultWriter.InsertCell();
                                            ResultWriter.Write("Duplicate Documents");
                                        }
                                        ResultWriter.EndRow();
                                    }
                                }
                            }
                            ResultWriter.EndTable();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("Error while applying license: " + ex.Message, LogFilePath);
                    }
                }

                MemoryStream UpdateDoc = new MemoryStream();
                if (Logging)
                    Log("Saving Document", LogFilePath);

                Result.Save(UpdateDoc, SaveFormat.Docx);
                byte[] byteData = UpdateDoc.ToArray();

                // Encode the data using base64.
                string encodedData = System.Convert.ToBase64String(byteData);

                if (Logging)
                    Log("Creating Attachment for result", LogFilePath);

                Entity NewNote = new Entity("annotation");
                // add Note to entity
                NewNote.Attributes.Add("objectid", new EntityReference(RecordType, ThisRecordId));
                NewNote.Attributes.Add("subject", "Duplicate detection report");

                // Set EncodedData to Document Body
                NewNote.Attributes.Add("documentbody", encodedData);

                // Set the type of attachment
                NewNote.Attributes.Add("mimetype", @"application\ms-word");
                NewNote.Attributes.Add("notetext", "Duplicate detection report");

                // Set the File Name
                NewNote.Attributes.Add("filename", "Duplicate detection report");

                Guid NewNoteId = service.Create(NewNote);
                OutputAttachmentId.Set(executionContext, new EntityReference("annotation", NewNoteId));

                if (Logging)
                    Log("Attachment Created Successfully", LogFilePath);

            }
            else if (detectIn == 1) //under this entity
            {
                Guid ThisRecordId = context.PrimaryEntityId;
                string RecordType = context.PrimaryEntityName;
                Document Result = new Document();
                DocumentBuilder ResultWriter = new DocumentBuilder(Result);

                if (Logging)
                    Log("Working under all attachments under this Entity", LogFilePath);
                QueryExpression RetrieveNoteQuery = new QueryExpression("annotation");
                RetrieveNoteQuery.ColumnSet = new ColumnSet(new string[] { "filename", "subject", "documentbody", "objectid" });
                if (Logging)
                    Log("Executing Query to retrieve All Notes within this Entity", LogFilePath);
                EntityCollection Notes = service.RetrieveMultiple(RetrieveNoteQuery);

                foreach (Entity Note in Notes.Entities)
                {
                    if (Note.Contains("objectid") && ((EntityReference)Note["objectid"]).LogicalName == RecordType)
                    {
                        try
                        {
                            if (Note.Contains("documentbody"))
                            {
                                string FileName = "";
                                if (Note.Contains("filename"))
                                    FileName = Note["filename"].ToString();

                                byte[] DocumentBody = Convert.FromBase64String(Note["documentbody"].ToString());
                                MemoryStream fileStream = new MemoryStream(DocumentBody);
                                Document doc = new Document(fileStream);

                                ResultWriter.Writeln("Comparing Document: " + FileName);
                                ResultWriter.StartTable();

                                foreach (Entity OtherNote in Notes.Entities)
                                {
                                    if (OtherNote.Id != Note.Id)
                                    {
                                        if (OtherNote.Contains("documentbody"))
                                        {
                                            string OtherFileName = "";
                                            if (OtherNote.Contains("filename"))
                                                OtherFileName = OtherNote["filename"].ToString();
                                            byte[] OtherDocumentBody = Convert.FromBase64String(OtherNote["documentbody"].ToString());
                                            MemoryStream fileStream2 = new MemoryStream(OtherDocumentBody);
                                            Document doc2 = new Document(fileStream);

                                            ResultWriter.InsertCell();
                                            ResultWriter.Write(OtherFileName);

                                            doc.Compare(doc2, "a", DateTime.Now);
                                            if (doc.Revisions.Count == 0)
                                            {
                                                ResultWriter.InsertCell();
                                                ResultWriter.Write("Duplicate Documents");
                                            }
                                            ResultWriter.EndRow();
                                        }
                                    }
                                }
                                ResultWriter.EndTable();
                            }
                        }
                        catch (Exception ex)
                        {
                            Log("Error while applying license: " + ex.Message, LogFilePath);
                        }
                    }
                }
                MemoryStream UpdateDoc = new MemoryStream();
                if (Logging)
                    Log("Saving Document", LogFilePath);

                Result.Save(UpdateDoc, SaveFormat.Docx);
                byte[] byteData = UpdateDoc.ToArray();

                // Encode the data using base64.
                string encodedData = System.Convert.ToBase64String(byteData);

                if (Logging)
                    Log("Creating Attachment for result", LogFilePath);

                Entity NewNote = new Entity("annotation");
                // add Note to entity
                NewNote.Attributes.Add("objectid", new EntityReference(RecordType, ThisRecordId));
                NewNote.Attributes.Add("subject", "Duplicate detection report");

                // Set EncodedData to Document Body
                NewNote.Attributes.Add("documentbody", encodedData);

                // Set the type of attachment
                NewNote.Attributes.Add("mimetype", @"application\ms-word");
                NewNote.Attributes.Add("notetext", "Duplicate detection report");

                // Set the File Name
                NewNote.Attributes.Add("filename", "Duplicate detection report");

                Guid NewNoteId = service.Create(NewNote);
                OutputAttachmentId.Set(executionContext, new EntityReference("annotation", NewNoteId));

                if (Logging)
                    Log("Attachment Created Successfully", LogFilePath);

            }
            else if (detectIn == 2)//under whole organization
            {
                Guid ThisRecordId = context.PrimaryEntityId;
                string RecordType = context.PrimaryEntityName;
                Document Result = new Document();
                DocumentBuilder ResultWriter = new DocumentBuilder(Result);

                if (Logging)
                    Log("Working under all attachments under this Entity", LogFilePath);
                QueryExpression RetrieveNoteQuery = new QueryExpression("annotation");
                if (Logging)
                    Log("Executing Query to retrieve All Notes within this Entity", LogFilePath);
                EntityCollection Notes = service.RetrieveMultiple(RetrieveNoteQuery);

                foreach (Entity Note in Notes.Entities)
                {
                    try
                    {
                        if (Note.Contains("documentbody"))
                        {
                            string FileName = "";
                            if (Note.Contains("filename"))
                                FileName = Note["filename"].ToString();

                            byte[] DocumentBody = Convert.FromBase64String(Note["documentbody"].ToString());
                            MemoryStream fileStream = new MemoryStream(DocumentBody);
                            Document doc = new Document(fileStream);

                            ResultWriter.Writeln("Comparing Document: " + FileName);
                            ResultWriter.StartTable();

                            foreach (Entity OtherNote in Notes.Entities)
                            {
                                if (OtherNote.Id != Note.Id)
                                {
                                    if (OtherNote.Contains("documentbody"))
                                    {
                                        string OtherFileName = "";
                                        if (OtherNote.Contains("filename"))
                                            OtherFileName = OtherNote["filename"].ToString();
                                        byte[] OtherDocumentBody = Convert.FromBase64String(OtherNote["documentbody"].ToString());
                                        MemoryStream fileStream2 = new MemoryStream(OtherDocumentBody);
                                        Document doc2 = new Document(fileStream);

                                        ResultWriter.InsertCell();
                                        ResultWriter.Write(OtherFileName);

                                        doc.Compare(doc2, "a", DateTime.Now);
                                        if (doc.Revisions.Count == 0)
                                        {
                                            ResultWriter.InsertCell();
                                            ResultWriter.Write("Duplicate Documents");
                                        }
                                        ResultWriter.EndRow();
                                    }
                                }
                            }
                            ResultWriter.EndTable();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("Error while applying license: " + ex.Message, LogFilePath);
                    }

                }
                MemoryStream UpdateDoc = new MemoryStream();
                if (Logging)
                    Log("Saving Document", LogFilePath);

                Result.Save(UpdateDoc, SaveFormat.Docx);
                byte[] byteData = UpdateDoc.ToArray();

                // Encode the data using base64.
                string encodedData = System.Convert.ToBase64String(byteData);

                if (Logging)
                    Log("Creating Attachment for result", LogFilePath);

                Entity NewNote = new Entity("annotation");
                // add Note to entity
                NewNote.Attributes.Add("objectid", new EntityReference(RecordType, ThisRecordId));
                NewNote.Attributes.Add("subject", "Duplicate detection report");

                // Set EncodedData to Document Body
                NewNote.Attributes.Add("documentbody", encodedData);

                // Set the type of attachment
                NewNote.Attributes.Add("mimetype", @"application\ms-word");
                NewNote.Attributes.Add("notetext", "Duplicate detection report");

                // Set the File Name
                NewNote.Attributes.Add("filename", "Duplicate detection report");

                Guid NewNoteId = service.Create(NewNote);
                OutputAttachmentId.Set(executionContext, new EntityReference("annotation", NewNoteId));

                if (Logging)
                    Log("Attachment Created Successfully", LogFilePath);

            }

        }
        private void Log(string Message, string LogFilePath)
        {
            try
            {
                if (LogFilePath == "")
                    File.AppendAllText("C:\\Aspose Logs\\Aspose.DuplicateDetector.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
                else
                    File.AppendAllText(LogFilePath + "\\Aspose.DuplicateDetector.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
