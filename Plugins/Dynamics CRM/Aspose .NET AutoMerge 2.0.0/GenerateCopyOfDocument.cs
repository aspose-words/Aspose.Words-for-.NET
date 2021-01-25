using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.AutoMerge
{
    public class GenerateCopyOfDocument : CodeActivity
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
        [Input("Attachment")]
        [ReferenceTarget("annotation")]
        public InArgument<EntityReference> AttachmentId { get; set; }

        [Output("Attachment")]
        [ReferenceTarget("annotation")]
        public OutArgument<EntityReference> OutputAttachmentId { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            Boolean Logging = EnableLogging.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
            EntityReference Attachment = AttachmentId.Get(executionContext);
            OutputAttachmentId.Set(executionContext, new EntityReference("annotation", Guid.Empty));
            try
            {
                if (Logging)
                    Log("Workflow Execution Start", LogFilePath);
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                if (Logging)
                    Log("Reading Attachment", LogFilePath);
                Entity ExistingAttachment = service.Retrieve(Attachment.LogicalName, Attachment.Id, new ColumnSet(true));
                if (ExistingAttachment != null)
                {
                    if (Logging)
                        Log("Creating New Attachment", LogFilePath);

                    // Create new Attachment under Email Activity.
                    Entity NewAttachment = new Entity("annotation");
                    if (ExistingAttachment.Contains("subject"))
                        NewAttachment.Attributes.Add("subject", ExistingAttachment["subject"]);
                    if (ExistingAttachment.Contains("filename"))
                        NewAttachment.Attributes.Add("filename", ExistingAttachment["filename"]);
                    if (ExistingAttachment.Contains("mimetype"))
                        NewAttachment.Attributes.Add("mimetype", ExistingAttachment["mimetype"]);
                    if (ExistingAttachment.Contains("documentbody"))
                        NewAttachment.Attributes.Add("documentbody", ExistingAttachment["documentbody"]);
                    Guid NewAttachmentId = service.Create(NewAttachment);
                    OutputAttachmentId.Set(executionContext, new EntityReference("annotation", NewAttachmentId));
                    if (Logging)
                        Log("New Attachment Created", LogFilePath);
                }
                else
                {
                    if (Logging)
                        Log("Provided Attachment doesnot exist", LogFilePath);
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
                    File.AppendAllText("C:\\Aspose Logs\\Aspose.AutoMerge.GenerateCopyOfDocument.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
                else
                    File.AppendAllText(LogFilePath + "\\Aspose.AutoMerge.GenerateCopyOfDocument.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            }
            catch { }
        }
    }
}
