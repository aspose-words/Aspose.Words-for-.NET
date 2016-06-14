using System.IO;
using Aspose.Words;
using System;
using System.Drawing;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertTCField
    {
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertTCField
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Initialize document.
            Document doc = new Document();

            // Create a document builder to insert content with.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TC field at the current document builder position.
            builder.InsertField("TC \"Entry Text\" \\f t");

            dataDir = dataDir + "DocumentBuilderInsertTCField_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertTCField
            Console.WriteLine("\nTC field inserted successfully into a document.\nFile saved at " + dataDir);
        }     
    }
}
