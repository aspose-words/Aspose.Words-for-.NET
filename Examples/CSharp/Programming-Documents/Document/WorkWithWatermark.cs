using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WorkWithWatermark
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            AddTextWatermarkWithSpecificOptions(dataDir);
            AddImageWatermarkWithSpecificOptions(dataDir);
            RemoveWatermarkFromDocument(dataDir);
        }

        public static void AddTextWatermarkWithSpecificOptions(string dataDir)
        {
            // ExStart: AddTextWatermarkWithSpecificOptions
            Document doc = new Document(dataDir + "Document.doc");

            TextWatermarkOptions options = new TextWatermarkOptions()
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Black,
                Layout = WatermarkLayout.Horizontal,
                IsSemitrasparent = false
            };

            doc.Watermark.SetText("Test", options);

            doc.Save(dataDir + "AddTextWatermark_out.docx");
            // ExEnd: AddTextWatermarkWithSpecificOptions
            Console.WriteLine("\nDocument saved successfully.\nFile saved at " + dataDir);
        }

        public static void AddImageWatermarkWithSpecificOptions(string dataDir)
        {
            // ExStart: AddImageWatermarkWithSpecificOptions
            Document doc = new Document(dataDir + "Document.doc");

            ImageWatermarkOptions options = new ImageWatermarkOptions()
            {
                Scale = 5,
                IsWashout = false
            };

            doc.Watermark.SetImage(Image.FromFile(dataDir + "Watermark.png"), options);

            doc.Save(dataDir + "AddImageWatermark_out.docx");
            // ExEnd: AddImageWatermarkWithSpecificOptions
            Console.WriteLine("\nDocument saved successfully.\nFile saved at " + dataDir);
        }

        public static void RemoveWatermarkFromDocument(string dataDir)
        {
            // ExStart: RemoveWatermarkFromDocument
            Document doc = new Document(dataDir + "AddTextWatermark_out.docx");

            if (doc.Watermark.Type == WatermarkType.Text)
            {
                doc.Watermark.Remove();
            }

            doc.Save(dataDir + "RemoveWatermark_out.docx");
            // ExEnd: RemoveWatermarkFromDocument
            Console.WriteLine("\nDocument saved successfully.\nFile saved at " + dataDir);
        }
    }
}
