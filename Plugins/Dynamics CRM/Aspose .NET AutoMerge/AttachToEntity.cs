using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.AutoMerge
{
    public class AttachToEntity : CodeActivity
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

        [RequiredArgument]
        [Input("Entity Logical Name")]
        public InArgument<string> EntityName { get; set; }

        [RequiredArgument]
        [Input("Record ID")]
        public InArgument<string> RecordId { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            Boolean Logging = EnableLogging.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
            EntityReference Attachment = AttachmentId.Get(executionContext);
            string entityName = EntityName.Get(executionContext);
            string recordId = RecordId.Get(executionContext);
            try
            {
                if (Logging)
                    Log("Workflow Execution Start", LogFilePath);
                if (ValidateParameters(executionContext))
                {
                    // Create a CRM Service in Workflow.
                    IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                    IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    if (Logging)
                        Log("Attaching Attachment", LogFilePath);

                    // Create an attachment.
                    Entity UpdatedAttachment = new Entity("annotation");
                    UpdatedAttachment.Id = Attachment.Id;
                    UpdatedAttachment.Attributes.Add("objectid", new EntityReference(entityName, new Guid(recordId)));
                    service.Update(UpdatedAttachment);

                    if (Logging)
                        Log("Attachment linked successfully", LogFilePath);

                    if (Logging)
                        Log("Workflow Executed Successfully", LogFilePath);
                }
            }
            catch (Exception ex)
            {
                Log(ex.Message, LogFilePath);
            }
        }

        private bool ValidateParameters(CodeActivityContext executionContext)
        {
            Boolean Logging = EnableLogging.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
            EntityReference Attachment = AttachmentId.Get(executionContext);
            string entityName = EntityName.Get(executionContext);
            string recordId = RecordId.Get(executionContext);
            return true;
        }

        private void Log(string Message, string LogFilePath)
        {
            try
            {
                if (LogFilePath == "")
                    File.AppendAllText("C:\\Aspose Logs\\Aspose.AutoMerge.AttachToAnyEntity.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
                else
                    File.AppendAllText(LogFilePath + "\\Aspose.AutoMerge.AttachToAnyEntity.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            }
            catch { }
        }
    }
}
