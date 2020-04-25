using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Web.Http;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using Font = System.Drawing.Font;

namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeWordsWatermark class to add or remove watermark from word file
	///</Summary>
	public class AsposeWordsWatermark : AsposeWordsBase
	{
		///<Summary>
		/// TextWatermark method to add text watermark
		///</Summary>

		public Response TextWatermark(string fileName, string folderName, string watermarkText, string watermarkColor,
		  string fontFamily = "Arial", double fontSize = 72, double textAngle = -45)
		{
			Opts.AppName = "Watermark";
			Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
			Opts.FolderName = folderName;
			Opts.FileName = fileName;
			Opts.OutputType = "docx";
			Opts.ResultFileName = Path.GetFileNameWithoutExtension(fileName) + " Text Watermark";

			return Process((inFilePath, outPath, zipOutFolder) =>
		   {
			   var doc = new Document(Opts.WorkingFileName);

			   if (string.IsNullOrEmpty(watermarkColor))
				   watermarkColor = "#FF808080"; // Gray
		  var color = ColorTranslator.FromHtml(watermarkColor.StartsWith("#") ? watermarkColor : "#" + watermarkColor);

			   var builder = new DocumentBuilder(doc);
			   var watermark = builder.InsertShape(ShapeType.TextBox, 1, 1); // DML
		  var run = new Run(doc, watermarkText)
			   {
				   Font = { Name = fontFamily, Color = color, Size = fontSize }
			   };
			   var paragraph = new Paragraph(doc);
			   paragraph.AppendChild(run);
			   watermark.AppendChild(paragraph);
			   watermark.TextBox.FitShapeToText = true;
			   watermark.TextBox.TextBoxWrapMode = TextBoxWrapMode.None;
			   watermark.TextBox.InternalMarginBottom = 0;
			   watermark.TextBox.InternalMarginLeft = 0;
			   watermark.TextBox.InternalMarginRight = 0;
			   watermark.TextBox.InternalMarginTop = 0;
			   watermark.Rotation = textAngle;
			   watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
			   watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
			   watermark.WrapType = WrapType.None;
			   watermark.VerticalAlignment = VerticalAlignment.Center;
			   watermark.HorizontalAlignment = HorizontalAlignment.Center;
			   watermark.ZOrder = -10000; // Appear behind other images
		  watermark.BehindText = true;
			   watermark.IsLayoutInCell = false;
			   watermark.Stroked = false;
			   watermark.Name = "WaterMark";

			   AddWatermark(doc, watermark);

		  // Textbox must be saved as DML shape to enable rotation. So, OOXML compliance
		  // must be "Transitional" or "Strict".
		  var so = new OoxmlSaveOptions(SaveFormat.Docx) { Compliance = OoxmlCompliance.Iso29500_2008_Transitional };
			   doc.Save(outPath, so);
		   });
		}
		///<Summary>
		/// ImageWatermark method to add image watermark
		///</Summary>

		public Response ImageWatermark(string fileName, string folderName, string imageFileName, string imageFolderName,
		  bool grayScale = false, double zoom = 100, double imageAngle = -45)
		{
			Opts.AppName = "Watermark";
			Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
			Opts.FolderName = folderName;
			Opts.FileName = fileName;
			Opts.OutputType = "docx";
			Opts.ResultFileName = Path.GetFileNameWithoutExtension(fileName) + " Image Watermark";

			return Process((inFilePath, outPath, zipOutFolder) =>
		   {
			   var doc = new Document(Opts.WorkingFileName);
			   var watermark = new Shape(doc, ShapeType.Rectangle)
			   {
				   Rotation = imageAngle,
				   RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
				   RelativeVerticalPosition = RelativeVerticalPosition.Page,
				   WrapType = WrapType.None,
				   VerticalAlignment = VerticalAlignment.Center,
				   HorizontalAlignment = HorizontalAlignment.Center,
				   ZOrder = -10000, // Appear behind other images
			  BehindText = true,
				   Name = "WaterMark"
			   };

			   var imagefilename = Config.Configuration.WorkingDirectory + imageFolderName + "/" + imageFileName;
			   if (!System.IO.File.Exists(imagefilename))
				   imagefilename = Config.Configuration.OutputDirectory + imageFolderName + "/" + imageFileName;

			   var image = Image.FromFile(imagefilename);
			   watermark.ImageData.SetImage(ResizeImage(image, zoom / 100));
			   watermark.ImageData.GrayScale = grayScale;

			   AddWatermark(doc, watermark);
			   doc.Save(outPath);
		   });
		}

		private static void AddWatermark(Document doc, Shape watermark)
		{
			var watermarkPara = new Paragraph(doc);
			watermarkPara.AppendChild(watermark);
			foreach (Section sect in doc.Sections)
			{
				// There could be up to three different headers in each section, since we want
				// The watermark to appear on all pages, insert into all headers.
				InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary);
				InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst);
				InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven);
			}
		}

		private static void InsertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, HeaderFooterType headerType)
		{
			var header = sect.HeadersFooters[headerType];
			if (header == null)
			{
				// There is no header of the specified type in the current section, create it.
				header = new HeaderFooter(sect.Document, headerType);
				sect.HeadersFooters.Add(header);
			}
			header.AppendChild(watermarkPara.Clone(true));
		}
		///<Summary>
		/// RemoveWatermark method to remove watermark
		///</Summary>

		public Response RemoveWatermark(string fileName, string folderName)
		{
			Opts.AppName = "Watermark";
			Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
			Opts.FolderName = folderName;
			Opts.FileName = fileName;
			Opts.ResultFileName = Path.GetFileNameWithoutExtension(fileName) + " Removed Watermark";

			return Process((inFilePath, outPath, zipOutFolder) =>
		   {
			   var doc = new Document(Opts.WorkingFileName);
		  // In MS Word there is no any special property for watermark. However, watermarks are shapes with the name containing "WaterMark"
		  // See: https://forum.aspose.com/t/detect-a-watermark-on-word-document/44883/5
		  foreach (HeaderFooter hf in doc.GetChildNodes(NodeType.HeaderFooter, true))
				   foreach (Shape shape in hf.GetChildNodes(NodeType.Shape, true))
					   if (shape.Name.Contains("WaterMark") || shape.TextPath.Text.Contains("WaterMark")) // WORDSNET-15559
					  shape.Remove();
			   doc.Save(outPath);
		   });
		}

	}
}
