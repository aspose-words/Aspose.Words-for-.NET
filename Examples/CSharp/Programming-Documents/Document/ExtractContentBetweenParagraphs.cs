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

namespace CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenParagraphs
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_WorkingWithDocument();

            Document doc = new Document(dataDir + "TestFile.doc");

            // Gather the nodes. The GetChild method uses 0-based index
            Paragraph startPara = (Paragraph)doc.FirstSection.GetChild(NodeType.Paragraph, 6, true);
            Paragraph endPara = (Paragraph)doc.FirstSection.GetChild(NodeType.Paragraph, 10, true);
            // Extract the content between these nodes in the document. Include these markers in the extraction.
            ArrayList extractedNodes = Common.ExtractContent(startPara, endPara, true);

            // Insert the content into a new separate document and save it to disk.
            Document dstDoc = Common.GenerateDocument(doc, extractedNodes);
            dstDoc.Save(dataDir + "TestFile.Paragraphs Out.doc");

            Console.WriteLine("\nExtracted content betweenn the paragraphs successfully.\nFile saved at " + dataDir + "TestFile.Paragraphs Out.doc");
        }
    }
}
