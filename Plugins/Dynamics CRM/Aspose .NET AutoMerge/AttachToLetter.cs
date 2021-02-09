using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.AutoMerge
{
    public class AttachToLetter : CodeActivity
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
        [Input("Letter")]
        [ReferenceTarget("letter")]
        public InArgument<EntityReference> LetterId { get; set; }

        [RequiredArgument]
        [Input("Attachment")]
        [ReferenceTarget("annotation")]
        public InArgument<EntityReference> AttachmentId { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            Boolean Logging = EnableLogging.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
            EntityReference Letter = LetterId.Get(executionContext);
            EntityReference Attachment = AttachmentId.Get(executionContext);
            try
            {
                if (Logging)
                    Log("Workflow Execution Start", LogFilePath);

                // Create a CRM Service in Workflow.
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                if (Logging)
                    Log("Retrieving Attachment", LogFilePath);

                Entity TempAttachment = service.Retrieve("annotation", Attachment.Id, new ColumnSet(true));
                if (TempAttachment != null)
                {
                    if (Logging)
                        Log("Creating New Attachment", LogFilePath);

                    // Create an attachment.
                    Entity NewAttachment = new Entity("annotation");
                    if (TempAttachment.Contains("subject"))
                        NewAttachment.Attributes.Add("subject", TempAttachment["subject"]);
                    if (TempAttachment.Contains("filename"))
                        NewAttachment.Attributes.Add("filename", TempAttachment["filename"]);
                    if (TempAttachment.Contains("notetext"))
                        NewAttachment.Attributes.Add("notetext", TempAttachment["notetext"]);
                    if (TempAttachment.Contains("mimetype"))
                        NewAttachment.Attributes.Add("mimetype", TempAttachment["mimetype"]);
                    if (TempAttachment.Contains("documentbody"))
                        NewAttachment.Attributes.Add("documentbody", TempAttachment["documentbody"]);
                    NewAttachment.Attributes.Add("objectid", new EntityReference(Letter.LogicalName, Letter.Id));
                    service.Create(NewAttachment);

                    if (Logging)
                        Log("New Attachment Added To Letter", LogFilePath);
                }
                else
                {
                    if (Logging)
                        Log("Temp Attachment does not exist", LogFilePath);
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
            try
            {
                if (LogFilePath == "")
                    File.AppendAllText("C:\\Aspose Logs\\Aspose.AutoMerge.AttachToLetter.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
                else
                    File.AppendAllText(LogFilePath + "\\Aspose.AutoMerge.AttachToLetter.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            }
            catch { }
        }
    }
}
