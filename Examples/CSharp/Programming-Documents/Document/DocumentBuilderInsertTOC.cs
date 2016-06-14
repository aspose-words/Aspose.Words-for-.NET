using System.IO;
using Aspose.Words;
using System;
using System.Drawing;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertTOC
    {
        public static void Run()
        {
            //ExStart:DocumentBuilderInsertTOC
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Initialize document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // The newly inserted table of contents will be initially empty.
            // It needs to be populated by updating the fields in the document.
            //ExStart:UpdateFields
            doc.UpdateFields();
            //ExEnd:UpdateFields
            dataDir = dataDir + "DocumentBuilderInsertTOC_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertTOC
            Console.WriteLine("\nTable of contents field inserted successfully into a document.\nFile saved at " + dataDir);
        }     
    }
}
