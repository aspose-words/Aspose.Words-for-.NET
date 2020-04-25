using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace Aspose.Words.Live.Demos.UI.Controllers
{
	///<Summary>
	/// AsposeWordsSignatureController class to sign word document
	///</Summary>
	public class AsposeWordsSignatureController : AsposeWordsBase
  {
    private const int ImageHeight = 100;
    private const int RightPadding = 10;
    private const int TopPadding = 10;

		///<Summary>
		/// Sign documents
		///</Summary>
		[MimeMultipart]
		[HttpPost]
		[AcceptVerbs("GET", "POST")]
		public async Task<Response> Sign(string outputType, string signatureType)
		{
			try
			{
				var files = await UploadFiles();
				var docs = files.Where(x => x.Headers.ContentDisposition.Name != "\"imageFile\"").Select(x => new Document(x.LocalFileName)).ToArray();
        if (docs.Length == 0 || docs.Length > MaximumUploadFiles)
          return MaximumFileLimitsResponse;

        SetDefaultOptions(docs, outputType);
        Opts.AppName = "Signature";
        Opts.MethodName = "Sign";
        Opts.ResultFileName = Path.GetFileNameWithoutExtension(docs[0].OriginalFileName);
        Opts.ZipFileName = "Signed documents";

        var imageFile = files.FirstOrDefault(x => x.Headers.ContentDisposition.Name == "\"imageFile\"")?.LocalFileName;
        if (signatureType == "image" && imageFile == null)
          return new Response
          {
            Status = "Can't process the image file",
            StatusCode = 500
          };

        var form = HttpContext.Current.Request.Form;
				return  Process((inFilePath, outPath, zipOutFolder) =>
				{
					var tasks = docs.Select(doc => Task.Factory.StartNew(() =>
					{
						var builder = new DocumentBuilder(doc);
						builder.MoveToDocumentEnd();
						switch (signatureType)
						{
							case "drawing":
								AddDrawing(builder, form["image"]);
								break;
							case "text":
								AddText(builder, form["text"], form["textColor"]);
								break;
							case "image":
								AddImage(builder, imageFile);
								break;
						}
						doc.UpdatePageLayout();

						SaveDocument(doc, outPath, zipOutFolder);
					})).ToArray();

					Task.WaitAll(tasks);
				});
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response
				{
					Status = "Error during processing your file",
					StatusCode = 500
				};
			}
		}


		private void AddText(DocumentBuilder builder, string text, string textColor)
    {
      Color color;
      if (!string.IsNullOrEmpty(textColor))
      {
        if (!textColor.StartsWith("#"))
          textColor = "#" + textColor;
        color = ColorTranslator.FromHtml(textColor);
      }
      else
        color = Color.Black;

      var font = builder.Font;
      font.Size = 16;
      font.Bold = true;
      font.Color = color;
      font.Name = "Arial";

      var paragraph = builder.ParagraphFormat;
      paragraph.Alignment = ParagraphAlignment.Right;
      builder.Writeln(text);
    }

    private void AddDrawing(DocumentBuilder builder, string imageBinary)
    {
      var imageBytes = Convert.FromBase64String(imageBinary);
      
      using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
      {
        var bitmap = new Bitmap(ms);
        var width = ImageHeight * bitmap.Width / bitmap.Height;
        var shape = builder.InsertImage(bitmap, RelativeHorizontalPosition.RightMargin, -width - RightPadding,
          RelativeVerticalPosition.Paragraph, TopPadding, width, ImageHeight, WrapType.None);
        shape.BehindText = false;
      }
    }

    private void AddImage(DocumentBuilder builder, string imageFile)
    {
      var pageinfo = builder.Document.GetPageInfo(builder.Document.PageCount - 1);
      var maxwidth = pageinfo.WidthInPoints / 3;
      var maxheight = pageinfo.HeightInPoints / 5;
      var filename = imageFile;

      using (var bitmap = new Bitmap(filename))
      {
        float width = bitmap.Width;
        float height = bitmap.Height;
        if (width > maxwidth)
        {
          height *= maxwidth / width;
          width = maxwidth;
        }

        if (height > maxheight)
        {
          width *= maxheight / height;
          height = maxheight;
        }

        var shape = builder.InsertImage(bitmap, RelativeHorizontalPosition.RightMargin, -width - RightPadding,
          RelativeVerticalPosition.Paragraph, TopPadding, width, height, WrapType.None);
        shape.BehindText = false;
      }
    }

    private void AddImage(DocumentBuilder builder, string imageFoler, string imageFile)
    {
      AddImage(builder, Config.Configuration.WorkingDirectory + imageFoler + "/" + imageFile);
    }
  }
}
