//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Aspose.Words;
using Aspose.Words.Layout;

namespace PageNumbersOfNodesExample
{
    public class Program
    {
        public static void Main()
        {
            string dataDir = Path.GetFullPath("../../../Data/");

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

            Console.ReadLine();
        }
    }
}