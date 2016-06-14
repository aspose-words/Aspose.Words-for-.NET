using System.IO;
using Aspose.Words;
using System;
using System.Drawing;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class WriteAndFont
    {
        public static void Run()
        {
            //ExStart:WriteAndFont
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Initialize document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font formatting before adding text.
            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            builder.Write("Sample text.");           
            dataDir = dataDir + "WriteAndFont_out_.doc";
            doc.Save(dataDir);
            //ExEnd:WriteAndFont
            Console.WriteLine("\nFormatted text using DocumentBuilder inserted successfully.\nFile saved at " + dataDir);
        }        
    }
}
