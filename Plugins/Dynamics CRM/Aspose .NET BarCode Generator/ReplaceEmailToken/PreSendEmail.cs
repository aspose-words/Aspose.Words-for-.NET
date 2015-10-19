using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.BarCodeGenerator.ReplaceEmailToken
{
    public class PreSendEmail : IPlugin
    {
        private string LogFilePath = "C:\\Aspose Logs";
        public void Execute(IServiceProvider serviceProvider)
        {
            Log("Plugin Started");
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            //Creating Service Factory
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            //Creating Service
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            if (context.InputParameters.Contains("EmailId") && context.InputParameters["EmailId"] is Guid)
            {
                Guid PrimaryEntityId = (Guid)context.InputParameters["EmailId"];
                //Checking if target entity is not email, the code should return and not run.
                if (PrimaryEntityId == null || PrimaryEntityId == Guid.Empty)
                    return;
                try
                {
                    Log("Pre-Retrieve Email");
                    Entity Email = service.Retrieve("email", PrimaryEntityId, new ColumnSet(new string[] { "description" }));
                    Log("Post-Retrieve Email");
                    if (Email != null && Email.Contains("description"))
                    {
                        string EmailBody = Email["description"].ToString();
                        if (EmailBody.Contains("[AsposeBarCode{"))
                        {
                            int StartIndex = EmailBody.IndexOf("[AsposeBarCode{");
                            int EndIndex = EmailBody.IndexOf("}]", StartIndex);
                            string ConfigIdvalue = EmailBody.Substring(StartIndex + 15, EndIndex - (StartIndex + 15));
                            Log("Configuration ID:" + ConfigIdvalue);
                            Guid ConfigId = new Guid(ConfigIdvalue);
                            Entity BarCodeConfiguration = service.Retrieve("aspose_barcodeconfiguration", ConfigId, new ColumnSet(true));
                            if (BarCodeConfiguration != null)
                            {
                                //Get the Code text for the barcode
                                string CodeText = BarCodeConfiguration.Contains("aspose_barcodedata") ? BarCodeConfiguration["aspose_barcodedata"].ToString() : "Aspose .NET BarCode Data";
                                string SymbologyText = BarCodeConfiguration.Contains("aspose_symbology") ? BarCodeConfiguration.FormattedValues["aspose_symbology"].ToString() : "Code128";

                                string WebUrl = "http://199.83.230.209:1002/GenerateBarCode.aspx";
                                string querystring = "?symbology=" + SymbologyText + "&codetext=" + CodeText;
                                EmailBody = EmailBody.Replace("[AsposeBarCode{" + ConfigIdvalue + "}]", "<img alt='Aspose .NET BarCode Generator' src='" + WebUrl + querystring + "' />");

                                Entity UpdatedEmail = new Entity("email");
                                UpdatedEmail.Id = PrimaryEntityId;
                                UpdatedEmail.Attributes.Add("description", EmailBody);
                                service.Update(UpdatedEmail);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    //Log(ex.Message);
                    throw ex;
                }
            }
        }

        private void Log(string Message)
        {
            if (LogFilePath == "")
                File.AppendAllText("C:\\Aspose Logs\\Aspose.BarCodeGenerator.ReplaceEmailToken.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            else
                File.AppendAllText(LogFilePath + "\\Aspose.BarCodeGenerator.ReplaceEmailToken.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);

        }
    }
}
