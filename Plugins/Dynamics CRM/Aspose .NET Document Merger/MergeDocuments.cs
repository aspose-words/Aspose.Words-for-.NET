using Aspose.Words;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.DocumentMerger.MergeDocumentsInEntity
{
    public class MergeDocuments : CodeActivity
    {
        [Input("First Attachment")]
        [ReferenceTarget("annotation")]
        public InArgument<EntityReference> FirstAttachmentId { get; set; }

        [Input("Second Attachment")]
        [ReferenceTarget("annotation")]
        public InArgument<EntityReference> SecondAttachmentId { get; set; }

        [Input("Output")]
        public InArgument<string> OutputOption { get; set; }

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
                EntityReference FirstAttachment = FirstAttachmentId.Get(executionContext);
                EntityReference SecondAttachment = SecondAttachmentId.Get(executionContext);
                EntityReference Contact = ContactId.Get(executionContext);
                string Option = OutputOption.Get(executionContext);
                Boolean Logging = EnableLogging.Get(executionContext);
                string LicenseFilePath = LicenseFile.Get(executionContext);

                if (Logging)
                    Log("Workflow Executed");

                // Create a CRM Service in Workflow.
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                if (Logging)
                    Log("Executing Query to retrieve First Attachment");

                Entity FirstNote = service.Retrieve("annotation", FirstAttachment.Id, new ColumnSet(new string[] { "subject", "documentbody" }));

                if (Logging)
                    Log("First Attachment Retrieved Successfully");
                if (Logging)
                    Log("Executing Query to retrieve Second Attachment");

                Entity SecondNote = service.Retrieve("annotation", SecondAttachment.Id, new ColumnSet(new string[] { "subject", "documentbody" }));

                if (Logging)
                    Log("Second Attachment Retrieved Successfully");

                MemoryStream fileStream1 = null, fileStream2 = null;
                string FileName1 = "";
                string FileName2 = "";
                if (FirstNote != null && FirstNote.Contains("documentbody"))
                {
                    byte[] DocumentBody = Convert.FromBase64String(FirstNote["documentbody"].ToString());
                     fileStream1 = new MemoryStream(DocumentBody);
                     if (FirstNote.Contains("filename"))
                         FileName1 = FirstNote["filename"].ToString();
                }
                if (SecondNote != null && SecondNote.Contains("documentbody"))
                {
                    byte[] DocumentBody = Convert.FromBase64String(SecondNote["documentbody"].ToString());
                    fileStream2 = new MemoryStream(DocumentBody);
                    if (SecondNote.Contains("filename"))
                        FileName2 = SecondNote["filename"].ToString();
                }
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
                    Log("Merging Documents");

                Document doc1 = new Document(fileStream1);
                Document doc2 = new Document(fileStream2);
                doc1.AppendDocument(doc2, ImportFormatMode.KeepSourceFormatting);

                if (Logging)
                    Log("Merging Complete");

                MemoryStream UpdateDoc = new MemoryStream();

                if (Logging)
                    Log("Saving Document");

                doc1.Save(UpdateDoc, SaveFormat.Docx);
                byte[] byteData = UpdateDoc.ToArray();

                // Encode the data using base64.
                string encodedData = System.Convert.ToBase64String(byteData);

                if (Logging)
                    Log("Creating Attachment");

                Entity NewNote = new Entity("annotation");

                // Add a note to the entity.
                NewNote.Attributes.Add("objectid", new EntityReference("contact", Contact.Id));
                
                // Set EncodedData to Document Body.
                NewNote.Attributes.Add("documentbody", encodedData);

                // Set the type of attachment.
                NewNote.Attributes.Add("mimetype", @"application\ms-word");
                NewNote.Attributes.Add("notetext", "Document Created using template");

                if (Option == "0")
                {
                    NewNote.Id = FirstNote.Id;
                    NewNote.Attributes.Add("subject", FileName1);
                    NewNote.Attributes.Add("filename", FileName1);
                    service.Update(NewNote);
                }
                else if (Option == "1")
                {
                    NewNote.Id = SecondNote.Id;
                    NewNote.Attributes.Add("subject", FileName2);
                    NewNote.Attributes.Add("filename", FileName2);
                    service.Update(NewNote);
                }
                else
                {
                    NewNote.Attributes.Add("subject", "Aspose .Net Document Merger");
                    NewNote.Attributes.Add("filename", "Aspose .Net Document Merger.docx");
                    service.Create(NewNote);
                }
                if (Logging)
                    Log("Successfull");
            }
            catch (Exception ex)
            {
                Log(ex.Message);
            }
        }

        private void Log(string Message)
        {
            File.AppendAllText("C:\\Aspose Logs\\Aspose.MergeDocuments.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
        }
    }
}
