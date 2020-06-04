using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using Aspose.Words.Live.Demos.UI.Config;
using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.IO;
using System.Text;
using System.Web;
using File = System.IO.File;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
    ///<Summary>
    /// AsposeWordsEditorController class to edit word document
    ///</Summary>
    public class AsposeWordsEditorController : AsposeWordsBase
	{
        ///<Summary>
        /// GetHTML method to get HTML
        ///</Summary>
        public string GetHTML(string fileName, string folderName)
        {
            Opts.AppName = "Editor";
            Opts.FileName = fileName;
            Opts.FolderName = folderName;
            Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

            try
            {
                var doc = new Document(Opts.WorkingFileName);
                var so = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportFontsAsBase64 = true,
                    ExportImagesAsBase64 = true,
                    CssStyleSheetType = CssStyleSheetType.Inline,
                    Encoding = UTF8WithoutBom,
                    HtmlVersion = HtmlVersion.Html5
                };
                using (var stream = new MemoryStream())
                {
                    doc.Save(stream, so);
                    // FIX the overwlow of span-elements in some documents
                    return UTF8WithoutBom.GetString(stream.ToArray()).Replace("letter-spacing:-107374182.4pt;", "");
                }
            }
            catch (Exception ex)
            {
				Console.WriteLine(ex.Message);
                return null;
            }
        }
        ///<Summary>
        /// UpdateContents method to update contents
        ///</Summary>
        public Response UpdateContents(string fileName, string htmldata, string outputType)
        {
            Opts.AppName = "Editor";
            Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            outputType = outputType.ToLower();

            var foldername = Guid.NewGuid().ToString();
            var fn = Path.GetFileNameWithoutExtension(fileName) + outputType;
            var resultfile = Config.Configuration.OutputDirectory + foldername + "/" + fn;
            Directory.CreateDirectory(Path.GetDirectoryName(resultfile));

            try
            {
                switch (outputType)
                {
                    case ".html":
                        File.WriteAllText(resultfile, htmldata);
                        break;
                    default:
                        var lo = new HtmlLoadOptions()
                        {
                            LoadFormat = LoadFormat.Html,
                            Encoding = Encoding.UTF8
                        };
                        Document doc;
                        using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(htmldata)))
                            doc = new Document(stream, lo);
                        doc.Save(resultfile);
                        break;
                }
                return new Response()
                {
                    FileName = HttpUtility.UrlEncode(fn),
                    FolderName = foldername,
                    StatusCode = 200
                };
            }
            catch (Exception ex)
            {
				Console.WriteLine(ex.Message);
                return new Response()
                {
                    FileName = HttpUtility.UrlEncode(fn),
                    FolderName = foldername,
                    StatusCode = 500,
                    Status = ex.Message
                };
            }
        }
    }
}
