using Aspose.Words;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Mvc;

namespace Aspose.MVC.ExportToPDF.Controllers
{
    public class AsposeExporterController : Controller
    {
        // GET: /AsposeExporter/
        public void Index(string Format="pdf")
        {
            ApplyLicense();

            string baseUrl = Request.Url.Authority;

            if (Request.ServerVariables["HTTPS"] == "on")
            {
                baseUrl = "https://" + baseUrl;
            }
            else
            {
                baseUrl = "http://" + baseUrl;
            }

            string refUrl = Request.UrlReferrer.AbsoluteUri;
            string html = new WebClient().DownloadString(refUrl);

            // To make the relative image paths work, the head section must include the base URL.
            html = html.Replace("</head>", string.Format("<base href='{0}'></base></head>", baseUrl));

            MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(html));
            Document doc = new Document(stream);
            string fileName = System.Guid.NewGuid() + "." + Format;

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
