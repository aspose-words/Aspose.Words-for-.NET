using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenRuns
    {
        public static void Run()
        {
            //ExStart:ExtractContentBetweenRuns
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            // Retrieve a paragraph from the first section.
            Paragraph para = (Paragraph)doc.GetChild(NodeType.Paragraph, 7, true);

            // Use some runs for extraction.
            Run startRun = para.Runs[1];
            Run endRun = para.Runs[4];

            // Extract the content between these nodes in the document. Include these markers in the extraction.
            ArrayList extractedNodes = Common.ExtractContent(startRun, endRun, true);

            // Get the node from the list. There should only be one paragraph returned in the list.
            Node node = (Node)extractedNodes[0];
            // Print the text of this node to the console.
            Console.WriteLine(node.ToString(SaveFormat.Text));
            //ExEnd:ExtractContentBetweenRuns
            Console.WriteLine("\nExtracted content betweenn the runs successfully.");
        }
    }
}
