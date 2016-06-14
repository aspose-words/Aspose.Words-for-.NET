using System.IO;
using Aspose.Words;
using System;
using System.Drawing;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertParagraph
    {
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertParagraph
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Initialize document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font formatting
            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = System.Drawing.Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            // Specify paragraph formatting
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.FirstLineIndent = 8;
            paragraphFormat.Alignment = ParagraphAlignment.Justify;
            paragraphFormat.KeepTogether = true;

            builder.Writeln("A whole paragraph.");
            dataDir = dataDir + "DocumentBuilderInsertParagraph_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertParagraph
            Console.WriteLine("\nParagraph inserted successfully into the document using DocumentBuilder.\nFile saved at " + dataDir);
        }        
    }
}
