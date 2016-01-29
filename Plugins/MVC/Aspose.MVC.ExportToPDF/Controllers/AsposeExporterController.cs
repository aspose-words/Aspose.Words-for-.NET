using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Aspose.MVC.ExportToPDF.Controllers
{
    public class AsposeExporterController : Controller
    {
        //
        // GET: /AsposeExporter/

        public void Index(string Format="pdf")
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
            string fileName = System.Guid.NewGuid().ToString() + "." + Format;
            doc.Save(System.Web.HttpContext.Current.Response, fileName, ContentDisposition.Inline, null);

            System.Web.HttpContext.Current.Response.End();
        }

        public void pdf()
        {
            Index("pdf");
        }

        public void docx()
        {
            Index("docx");
        }

        public void doc()
        {
            Index("doc");
        }

        public void dot()
        {
            Index("dot");
        }

        public void dotx()
        {
            Index("dotx");
        }

        public void docm()
        {
            Index("docm");
        }

        public void dotm()
        {
            Index("dotm");
        }

        public void odt()
        {
            Index("odt");
        }

        public void ott()
        {
            Index("ott");
        }

        public void rtf()
        {
            Index("rtf");
        }

        public void txt()
        {
            Index("txt");
        }

    }
}
