using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.AutoMerge
{
    public class SendEmail : CodeActivity
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
        [Input("Email")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailId { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            Boolean Logging = EnableLogging.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
            EntityReference Email = EmailId.Get(executionContext);
            try
            {
                if (Logging)
                    Log("Workflow Execution Start", LogFilePath);
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                if (Logging)
                    Log("Sending Email", LogFilePath);
                SendEmailRequest sendEmailreq = new SendEmailRequest();
                sendEmailreq.EmailId = Email.Id;
                sendEmailreq.TrackingToken = "";
                sendEmailreq.IssueSend = true;
                SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailreq);
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
                    File.AppendAllText("C:\\Aspose Logs\\Aspose.AutoMerge.SendEmail.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
                else
                    File.AppendAllText(LogFilePath + "\\Aspose.AutoMerge.SendEmail.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            }
            catch { }
        }
    }
}
