using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentUsingField
    {
        public static void Run()
        {
            //ExStart:ExtractContentUsingField
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

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
            dataDir = dataDir +  RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);
            //ExEnd:ExtractContentUsingField
            Console.WriteLine("\nExtracted content using the Field successfully.\nFile saved at " + dataDir);
        }
    }
}
