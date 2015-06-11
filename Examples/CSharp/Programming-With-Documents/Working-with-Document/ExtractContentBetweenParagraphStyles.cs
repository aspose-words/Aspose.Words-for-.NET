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
    class ExtractContentBetweenParagraphStyles
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_WorkingWithDocument();

            Document doc = new Document(dataDir + "TestFile.doc");

            // Gather a list of the paragraphs using the respective heading styles.
            ArrayList parasStyleHeading1 = Common.ParagraphsByStyleName(doc, "Heading 1");
            ArrayList parasStyleHeading3 = Common.ParagraphsByStyleName(doc, "Heading 3");

            // Use the first instance of the paragraphs with those styles.
            Node startPara1 = (Node)parasStyleHeading1[0];
            Node endPara1 = (Node)parasStyleHeading3[0];

            // Extract the content between these nodes in the document. Don't include these markers in the extraction.
            ArrayList extractedNodes = Common.ExtractContent(startPara1, endPara1, false);

            // Insert the content into a new separate document and save it to disk.
            Document dstDoc = Common.GenerateDocument(doc, extractedNodes);
            dstDoc.Save(dataDir + "TestFile.Styles Out.doc");

            Console.WriteLine("\nExtracted content betweenn the paragraph styles successfully.\nFile saved at " + dataDir + "TestFile.Styles Out.doc");
        }
    }
}
