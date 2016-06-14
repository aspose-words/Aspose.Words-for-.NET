
using System;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Quick_Start
{
    class WorkingWithNodes
    {
        public static void Run()
        {
            // Create a new document.
            Document doc = new Document();

            // Creates and adds a paragraph node to the document.
            Paragraph para = new Paragraph(doc);

            // Typed access to the last section of the document.
            Section section = doc.LastSection;
            section.Body.AppendChild(para);

            // Next print the node type of one of the nodes in the document.
            NodeType nodeType = doc.FirstSection.Body.NodeType;

            Console.WriteLine("\nNodeType: " + Node.NodeTypeToString(nodeType));
        }
    }
}
