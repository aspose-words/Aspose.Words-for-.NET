using Aspose.Words.Markup;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_StructuredDocumentTag
{
    class SetContentControlColor
    {
        public static void Run()
        {
            // ExStart:SetContentControlColor
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithStructuredDocumentTag();

            Document doc = new Document(dataDir + "input.docx");
            StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            sdt.Color = Color.Red;

            dataDir = dataDir + "SetContentControlColor_out.docx";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetContentControlColor
            Console.WriteLine("\nSet the color of content control successfully.");
        }
    }
}
