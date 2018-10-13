using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertHorizontalRule
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            // ExStart:DocumentBuilderInsertHorizontalRule
            // Initialize document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Insert a horizontal rule shape into the document.");
            builder.InsertHorizontalRule();

            dataDir = dataDir + "DocumentBuilder.InsertHorizontalRule_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderInsertHorizontalRule
            Console.WriteLine("\nBookmark using DocumentBuilder inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
