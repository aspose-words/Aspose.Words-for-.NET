using Aspose.Words.Markup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_StructuredDocumentTag
{
    public class ClearContentsControl
    {
        public static void Run()
        {
            // ExStart:ClearContentsControl
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithStructuredDocumentTag();

            Document doc = new Document(dataDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Clear();

            dataDir = dataDir + "ClearContentsControl_out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:ClearContentsControl
            Console.WriteLine("\nClear the contents of content control successfully.");
        }
    }
}
