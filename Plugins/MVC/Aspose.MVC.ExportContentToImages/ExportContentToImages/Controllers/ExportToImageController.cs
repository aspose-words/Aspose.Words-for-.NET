using System;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Web.Mvc;
using Aspose.Words;
using System.Text;
using Aspose.Words.Saving;

namespace ExportContentToImages.Controllers
{
    public class ExportToImageController : Controller
    {
        // GET: /ExportToImage/
        public ActionResult Index(string Format="png")
        {
            // License this component using an Aspose.Words license file,
            // if one exists at this location in the local file system.
            string licenseFile = Server.MapPath("~/App_Data/Aspose.Words.lic");

            if (System.IO.File.Exists(licenseFile))
            {
                License license = new License();
                license.SetLicense(licenseFile);
            }

            if (Request.UrlReferrer == null) 
                return RedirectToAction("index", "Home");

            string refUrl = Request.UrlReferrer.AbsoluteUri;
            string html = new WebClient().DownloadString(refUrl);

            Document doc;

            using (MemoryStream memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                doc = new Document(memoryStream);
            }

            string imageFileName = "";

            if (doc.PageCount > 1)
                Directory.CreateDirectory(Server.MapPath("~/Images/" + "Zip"));

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            if (Format.Contains("png"))
            {
                options = new ImageSaveOptions(SaveFormat.Png);
                options.PageSet = new PageSet(0);
            }
            else if (Format.Contains("JPEG"))
            {
                options = new ImageSaveOptions(SaveFormat.Jpeg);
                options.PageSet = new PageSet(0);
            }
            else if (Format.Contains("TIFF"))
            {
                options = new ImageSaveOptions(SaveFormat.Tiff);
                options.PageSet = new PageSet(0);
            }
            else if (Format.Contains("bmp"))
            {
                options = new ImageSaveOptions(SaveFormat.Bmp);
                options.PageSet = new PageSet(0);
            }

            // Create an "Images" folder, and populate it with images that we will render the document into.
            string imageFolderPath = Server.MapPath("~/Images");

            if (!Directory.Exists(imageFolderPath))
                Directory.CreateDirectory(imageFolderPath);


            for (int i = 0; i < doc.PageCount; i++)
            {
                imageFileName = $"{i}_{Guid.NewGuid()}.{Format}";;

                if (Format.Contains("TIFF"))
                {
                    options.PageSet = PageSet.All;
                    doc.Save(Server.MapPath("~/Images/Zip/") + imageFileName, options);
                }
                else
                {
                    options.PageSet = new PageSet(i);

                    if (doc.PageCount > 1)
                        doc.Save(Server.MapPath("~/Images/Zip/") + imageFileName, options);
                    else
                        doc.Save(Server.MapPath("~/Images/Zip/") + imageFileName, options);
                }
            }

            // If a webpage is large enough for multiple images, and the output image type is not "Tiff",
            // then download them all in one Zip. A single Tiff file will contain
            // all the images, so we do not need a Zip archive in this case.
            if (doc.PageCount > 1 && !Format.Contains("TIFF"))
            {
                string ImagePath = Server.MapPath("~/Images/Zip/");
                string downloadDirectory = Server.MapPath("~/Images/");
                ZipFile.CreateFromDirectory(ImagePath, downloadDirectory + "OutputImages.zip");

                System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                response.ClearContent();
                response.Clear();
                response.ContentType = "application/zip";
                response.AddHeader("Content-Disposition", "attachment; filename=OutputImages.zip" + ";");
                response.TransmitFile(downloadDirectory + "OutputImages.zip");
                response.End();

                Directory.Delete(ImagePath, true);
                System.IO.File.Delete(downloadDirectory + "OutputImages.zip");
            }
            else
            {
                string filepath = Server.MapPath("~/Images/Zip/") + imageFileName;
                string downloadDirectory = Server.MapPath("~/Images/Zip/");
                System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;

                response.ClearContent();
                response.Clear();
                response.ContentType = "image/"+Format;
                response.AddHeader("Content-Disposition", "attachment; filename="+imageFileName + ";");
                response.TransmitFile(filepath);
                response.End();

                Directory.Delete(downloadDirectory, true);
            }

            return RedirectToAction("index", "Home");
        }
    }
}
