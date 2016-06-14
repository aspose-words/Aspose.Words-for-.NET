using System.IO;
using Aspose.Words;
using System;
using System.Drawing;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertBreak
    {
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertBreak
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Initialize document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 3.");
            dataDir = dataDir + "DocumentBuilderInsertBreak_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertBreak
            Console.WriteLine("\nPage breaks inserted into a document using DocumentBuilder.\nFile saved at " + dataDir);
        }     
    }
}
