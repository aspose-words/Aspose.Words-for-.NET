using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ParagraphStyleSeparator
    {
        public static void Run()
        {
            // ExStart:ParagraphStyleSeparator
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();

            // Initialize document.
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (paragraph.BreakIsStyleSeparator)
                {
                    Console.WriteLine("Separator Found!");
                }
            }
            // ExEnd:ParagraphStyleSeparator
        }
    }
}
