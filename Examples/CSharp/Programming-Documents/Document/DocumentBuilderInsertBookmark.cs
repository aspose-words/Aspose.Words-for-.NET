using System.IO;
using Aspose.Words;
using System;
using System.Drawing;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertBookmark
    {
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertBookmark
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Initialize document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("FineBookmark");
            builder.Writeln("This is just a fine bookmark.");
            builder.EndBookmark("FineBookmark");

            dataDir = dataDir + "DocumentBuilderInsertBookmark_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertBookmark
            Console.WriteLine("\nBookmark using DocumentBuilder inserted successfully.\nFile saved at " + dataDir);
        }     
    }
}
