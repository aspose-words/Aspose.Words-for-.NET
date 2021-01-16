using Aspose.Words;
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Mvc;

namespace Aspose.MVC.ExportToWord.Controllers
{
    public class ExportToWordController : Controller
    {
        // GET: /ExportToWord/
        public void Index()
        {
            ApplyLicense();

            string baseUrl = Request.Url.Authority;
            baseUrl = Request.ServerVariables["HTTPS"] == "on" ? "https://" + baseUrl : "http://" + baseUrl;
            
            string refUrl = Request.UrlReferrer.AbsoluteUri;
            string html = new WebClient().DownloadString(refUrl);

            // To make the relative image paths work, the head section must include the base URL.
            html = html.Replace("</head>", string.Format("<base href='{0}'></base></head>", baseUrl));

            MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(html));
            Document doc = new Document(stream);
            string fileName = Guid.NewGuid() + ".doc";
            doc.Save(System.Web.HttpContext.Current.Response, fileName, ContentDisposition.Inline, null);

            System.Web.HttpContext.Current.Response.End();
        }

        private void ApplyLicense()
        {
            string licenseFile = Server.MapPath("~/App_Data/Aspose.Words.lic");
            if (System.IO.File.Exists(licenseFile))
            {
                License license = new License();
                license.SetLicense(licenseFile);
            }
        }
    }
}
