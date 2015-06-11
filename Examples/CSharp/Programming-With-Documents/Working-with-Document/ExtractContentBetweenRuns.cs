//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;

namespace CSharp.Programming_With_Documents.Working_with_Document
{
    class ExtractContentBetweenRuns
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_WorkingWithDocument();

            Document doc = new Document(dataDir + "TestFile.doc");

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

            Console.WriteLine("\nExtracted content betweenn the runs successfully.");
        }
    }
}
