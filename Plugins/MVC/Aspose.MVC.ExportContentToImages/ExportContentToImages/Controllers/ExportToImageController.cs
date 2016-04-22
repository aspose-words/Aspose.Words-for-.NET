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
        //
        // GET: /ExportToImage/


        public ActionResult Index(string Format="png")
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

            //Null value Check  
            if (Request.UrlReferrer != null)
            {
                string refUrl = Request.UrlReferrer.AbsoluteUri;

                string html = new WebClient().DownloadString(refUrl);

                var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(html));

                Document doc = new Document(memoryStream);
              
                string fileName = "";

                if (doc.PageCount>1)
                {
                    Directory.CreateDirectory(Server.MapPath("~/Images/" + "Zip"));
                }

                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

                if (Format.Contains("png"))
                {
                     options = new ImageSaveOptions(SaveFormat.Png);
                    options.PageCount = 1;
                }
                else if (Format.Contains("JPEG"))
                {
                     options = new ImageSaveOptions(SaveFormat.Jpeg);
                    options.PageCount = 1;
                    
                }
                else if (Format.Contains("TIFF"))
                {
                     options = new ImageSaveOptions(SaveFormat.Tiff);
                    options.PageCount = 1;
                }

                else if (Format.Contains("bmp"))
                {
                     options = new ImageSaveOptions(SaveFormat.Bmp);
                    options.PageCount = 1;
                }


                // Check for Images folder 
                string path = Server.MapPath("~/Images");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }


                // Convert the html , get page count and save PNG's in Images folder
                for (int i = 0; i < doc.PageCount; i++)
                {
                    if (Format.Contains("TIFF"))
                    {
                        options.PageCount = doc.PageCount;
                        fileName = i + "_" + System.Guid.NewGuid().ToString() + "." + Format;
                        doc.Save(Server.MapPath("~/Images/Zip/") + fileName, options);
                        break;
                    }
                    else
                    {
                        options.PageIndex = i;
                        if (doc.PageCount > 1)
                        {
                            fileName = i + "_" + System.Guid.NewGuid().ToString() + "." + Format;
                            doc.Save(Server.MapPath("~/Images/Zip/") + fileName, options);
                        }
                        else
                        {
                            // webpage count is 1 
                            fileName = i + "_" + System.Guid.NewGuid().ToString() + "." + Format;
                            doc.Save(Server.MapPath("~/Images/Zip/") + fileName, options);
                        }
                    }
                }

                /* if webpage count is more then one download images as a Zip but if image type if TIFF                
                  dont download as a zip because Tiff already have all content in one Image 
                 */
                if (doc.PageCount > 1 && !Format.Contains("TIFF"))
                {
                    try
                    {
                        string ImagePath = Server.MapPath("~/Images/Zip/");
                        string downloadDirectory = Server.MapPath("~/Images/");
                        ZipFile.CreateFromDirectory(ImagePath, downloadDirectory + "OutputImages.zip");
                        // Prompts user to save file
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
                    catch (Exception Ex)
                    {

                    }
                }
                else
                {
                    string filepath = Server.MapPath("~/Images/Zip/") + fileName;
                    string downloadDirectory = Server.MapPath("~/Images/Zip/");
                    System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                    response.ClearContent();
                    response.Clear();
                    response.ContentType = "image/"+Format;
                    response.AddHeader("Content-Disposition", "attachment; filename="+fileName + ";");
                    response.TransmitFile(filepath);
                    response.End();
                    Directory.Delete(downloadDirectory, true);
                }
            }
            //set the view to your default View (in my case its home index view)
            return RedirectToAction("index", "Home");
        }



        public void jpg()
        {
            Index("jpeg");
        }

        public void png()
        {
            Index("png");
        }

        public void bmp()
        {
            Index("bmp");
        }
         public void tiff()
        {
            Index("tiff");
        }

    }
}
