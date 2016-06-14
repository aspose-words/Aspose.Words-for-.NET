using System;
using System.Collections;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenParagraphs
    {
        public static void Run()
        {
            //ExStart:ExtractContentBetweenParagraphs
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            // Gather the nodes. The GetChild method uses 0-based index
            Paragraph startPara = (Paragraph)doc.FirstSection.GetChild(NodeType.Paragraph, 6, true);
            Paragraph endPara = (Paragraph)doc.FirstSection.GetChild(NodeType.Paragraph, 10, true);
            // Extract the content between these nodes in the document. Include these markers in the extraction.
            ArrayList extractedNodes = Common.ExtractContent(startPara, endPara, true);

            // Insert the content into a new separate document and save it to disk.
            Document dstDoc = Common.GenerateDocument(doc, extractedNodes);
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);
            //ExEnd:ExtractContentBetweenParagraphs
            Console.WriteLine("\nExtracted content betweenn the paragraphs successfully.\nFile saved at " + dataDir);
        }
    }
}
