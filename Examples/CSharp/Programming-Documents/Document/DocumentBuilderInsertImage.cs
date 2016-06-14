using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertImage
    {
        public static void Run()
        {
            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            InsertInlineImage(dataDir);
            InsertFloatingImage(dataDir);
        }
        public static void InsertInlineImage(string dataDir)
        {
            //ExStart:DocumentBuilderInsertInlineImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(dataDir + "Watermark.png");
            dataDir = dataDir + "DocumentBuilderInsertInlineImage_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertInlineImage
            Console.WriteLine("\nInline image using DocumentBuilder inserted successfully.\nFile saved at " + dataDir);
        }
        public static void InsertFloatingImage(string dataDir)
        {
            //ExStart:DocumentBuilderInsertFloatingImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(dataDir + "Watermark.png",
                RelativeHorizontalPosition.Margin,
                100,
                RelativeVerticalPosition.Margin,
                100,
                200,
                100,
                WrapType.Square);
            dataDir = dataDir + "DocumentBuilderInsertFloatingImage_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertFloatingImage
            Console.WriteLine("\nInline image using DocumentBuilder inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
