using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.AutoMerge
{
    public class DeleteTempDocument : CodeActivity
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

        protected override void Execute(CodeActivityContext executionContext)
        {
            Boolean Logging = EnableLogging.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
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
                    Log("Removing Attachment", LogFilePath);

                service.Delete(Attachment.LogicalName, Attachment.Id);

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
                    File.AppendAllText("C:\\Aspose Logs\\Aspose.AutoMerge.DeleteTempDocument.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
                else
                    File.AppendAllText(LogFilePath + "\\Aspose.AutoMerge.DeleteTempDocument.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            }
            catch { }
        }
        
    }
}
