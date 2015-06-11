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
using Aspose.Words.Fields;

namespace CSharp.Programming_With_Documents.Working_with_Document
{
    class ExtractContentUsingField
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_WorkingWithDocument();

            Document doc = new Document(dataDir + "TestFile.doc");

            // Use a document builder to retrieve the field start of a merge field.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
            // We could also get FieldStarts of a field using GetChildNode method as in the other examples.
            builder.MoveToMergeField("Fullname", false, false);

            // The builder cursor should be positioned at the start of the field.
            FieldStart startField = (FieldStart)builder.CurrentNode;
            Paragraph endPara = (Paragraph)doc.FirstSection.GetChild(NodeType.Paragraph, 5, true);

            // Extract the content between these nodes in the document. Don't include these markers in the extraction.
            ArrayList extractedNodes = Common.ExtractContent(startField, endPara, false);

            // Insert the content into a new separate document and save it to disk.
            Document dstDoc = Common.GenerateDocument(doc, extractedNodes);
            dstDoc.Save(dataDir + "TestFile.Fields Out.doc");

            Console.WriteLine("\nExtracted content using the Field successfully.\nFile saved at " + dataDir + "TestFile.Fields Out.doc");
        }
    }
}
