using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Aspose.MVC.ExportToWord.Controllers
{
    public class ExportToWordController : Controller
    {
        //
        // GET: /ExportToWord/

        public void Index()
        {
            string baseURL = Request.Url.Authority;

            if (Request.ServerVariables["HTTPS"] == "on")
            {
                baseURL = "https://" + baseURL;
            }
            else
            {
                baseURL = "http://" + baseURL;
            }

            // Check for license and apply if exists
            string licenseFile = Server.MapPath("~/App_Data/Aspose.Words.lic");
            if (System.IO.File.Exists(licenseFile))
            {
                License license = new License();
                license.SetLicense(licenseFile);
            }

            string refURL = Request.UrlReferrer.AbsoluteUri;

            string html = new WebClient().DownloadString(refURL);

            // To make the relative image paths work, base URL must be included in head section
            html = html.Replace("</head>", string.Format("<base href='{0}'></base></head>", baseURL));

            MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(html));
            Document doc = new Document(stream);
            string fileName = System.Guid.NewGuid().ToString() + ".doc";
            doc.Save(System.Web.HttpContext.Current.Response, fileName, ContentDisposition.Inline, null);

            System.Web.HttpContext.Current.Response.End();
        }

    }
}
