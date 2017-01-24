using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class PageNumbersOfNodes
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            Document doc = new Document(dataDir + "TestFile.docx");

            // Create and attach collector before the document before page layout is built.
            LayoutCollector layoutCollector = new LayoutCollector(doc);

            // This will build layout model and collect necessary information.
            doc.UpdatePageLayout();

            // Print the details of each document node including the page numbers. 
            foreach (Node node in doc.FirstSection.Body.GetChildNodes(NodeType.Any, true))
            {
                Console.WriteLine(" --------- ");
                Console.WriteLine("NodeType:   " + Node.NodeTypeToString(node.NodeType));
                Console.WriteLine("Text:       \"" + node.ToString(SaveFormat.Text).Trim() + "\"");
                Console.WriteLine("Page Start: " + layoutCollector.GetStartPageIndex(node));
                Console.WriteLine("Page End:   " + layoutCollector.GetEndPageIndex(node));
                Console.WriteLine(" --------- ");
                Console.WriteLine();
            }

            // Detatch the collector from the document.
            layoutCollector.Document = null;

            Console.WriteLine("\nFound the page numbers of all nodes successfully.");
        }
    }
}
