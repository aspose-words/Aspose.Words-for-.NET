using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.IO;

namespace Aspose.AutoMerge
{
    public class SaveToCloudStorage : CodeActivity
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
        [Input("Product URI")]
        [Default(@"http://api.aspose.com/v1.1")]
        public InArgument<string> ProductUri { get; set; }

        [RequiredArgument]
        [Input("App SID")]
        public InArgument<string> AppSID { get; set; }

        [RequiredArgument]
        [Input("App Key")]
        public InArgument<string> AppKey { get; set; }

        [RequiredArgument]
        [Input("Attachment")]
        [ReferenceTarget("annotation")]
        public InArgument<EntityReference> AttachmentId { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            Boolean EnableLoggingValue = EnableLogging.Get(executionContext);
            string ProductUriValue = ProductUri.Get(executionContext);
            string AppSIDValue = AppSID.Get(executionContext);
            string AppKeyValue = AppKey.Get(executionContext);
            string LogFilePath = LogFile.Get(executionContext);
            EntityReference Attachment = AttachmentId.Get(executionContext);
            CloudAppConfig config = new CloudAppConfig();
            config.ProductUri = ProductUriValue;
            config.AppSID = AppSIDValue;
            config.AppKey = AppKeyValue;
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (EnableLoggingValue)
                    Log("WorkFlow Started", LogFilePath);
                if (EnableLoggingValue)
                    Log("Retrieving Attachment From CRM", LogFilePath);
                Entity ThisAttachment = service.Retrieve("annotation", Attachment.Id, new ColumnSet(new string[] { "filename", "documentbody", "mimetype" }));
                if (ThisAttachment != null)
                {
                    if (EnableLoggingValue)
                        Log("Attachment Retrieved Successfully", LogFilePath);
                    if (ThisAttachment.Contains("mimetype") && ThisAttachment.Contains("documentbody"))
                    {
                        string FileName = "Aspose .NET AutoMerge Attachment (" + DateTime.Now.ToString() + ").docx";
                        if (ThisAttachment.Contains("filename"))
                            FileName = ThisAttachment["filename"].ToString();
                        config.FileName = FileName;
                        byte[] DocumentBody = Convert.FromBase64String(ThisAttachment["documentbody"].ToString());
                        MemoryStream fileStream = new MemoryStream(DocumentBody);

                        if (EnableLoggingValue)
                            Log("Upload Attachment on Storage", LogFilePath);
                        UploadFileOnStorage(config, fileStream);
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex.Message, LogFilePath);
                throw ex;
            }
        }

        private void UploadFileOnStorage(CloudAppConfig Config, MemoryStream fileStream)
        {
            string URIRequest = Config.ProductUri + "/storage/file/" + Config.FileName;
            string URISigned = Sign(URIRequest, Config.AppSID, Config.AppKey);
            try
            {
                System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(URISigned);
                req.Method = "PUT";
                req.ContentType = "application/x-www-form-urlencoded";
                req.AllowWriteStreamBuffering = true;
                using (System.IO.Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(fileStream.ToArray(), 0, (int)fileStream.Length);
                }
                string statusCode = null;
                using (System.Net.HttpWebResponse response = (System.Net.HttpWebResponse)req.GetResponse())
                {
                    statusCode = response.StatusCode.ToString();
                }
            }
            catch (System.Net.WebException webex)
            {
                throw new Exception(webex.Message);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public Stream ProcessCommand(string strURI, string strHttpCommand)
        {
            try
            {
                Uri address = new Uri(strURI);
                System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(address);
                request.Method = strHttpCommand;
                request.ContentType = "application/json";

                request.ContentLength = 0;
                System.Net.HttpWebResponse response = (System.Net.HttpWebResponse)request.GetResponse();
                return response.GetResponseStream();
            }
            catch (System.Net.WebException webex)
            {
                throw new Exception(webex.Message);
            }
            catch (Exception Ex)
            {
                throw new Exception(Ex.Message);
            }
        }

        public string Sign(string URIRequest, string AppSIDValue, string AppKeyValue)
        {
            try
            {
                // Add appSID parameter.
                UriBuilder builder = new UriBuilder(URIRequest);
                if (builder.Query != null && builder.Query.Length > 1)
                    builder.Query = builder.Query.Substring(1) + "&appSID=" + AppSIDValue;
                else
                    builder.Query = "appSID=" + AppSIDValue;

                // Remove final slash here as it can be added automatically.
                builder.Path = builder.Path.TrimEnd('/');

                byte[] privateKey = System.Text.Encoding.UTF8.GetBytes(AppKeyValue);

                System.Security.Cryptography.HMACSHA1 algorithm = new System.Security.Cryptography.HMACSHA1(privateKey);

                byte[] sequence = System.Text.ASCIIEncoding.ASCII.GetBytes(builder.Uri.AbsoluteUri);
                byte[] hash = algorithm.ComputeHash(sequence);
                string signature = Convert.ToBase64String(hash);

                // Remove invalid symbols.
                signature = signature.TrimEnd('=');

                //signature = System.Web.HttpUtility.UrlEncode(signature);
                signature = System.Uri.EscapeDataString(signature);

                // Convert codes to upper case as they can be updated automatically.
                signature = System.Text.RegularExpressions.Regex.Replace(signature, "%[0-9a-f]{2}", e => e.Value.ToUpper());

                // Add the signature to query string.
                return string.Format("{0}&signature={1}", builder.Uri.AbsoluteUri, signature);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void Log(string Message, string LogFilePath)
        {
            try
            {
                if (LogFilePath == "")
                    File.AppendAllText("C:\\Aspose Logs\\Aspose.AutoMerge.SaveToCloud.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
                else
                    File.AppendAllText(LogFilePath + "\\Aspose.AutoMerge.SaveToCloud.log", Environment.NewLine + DateTime.Now.ToString() + ":- " + Message);
            }
            catch { }
        }
    }
    struct CloudAppConfig
    {
        public string ProductUri;
        public string FileName;
        public string AppSID;
        public string AppKey;
    }
}
