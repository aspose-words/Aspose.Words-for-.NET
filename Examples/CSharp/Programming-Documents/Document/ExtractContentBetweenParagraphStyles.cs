using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenParagraphStyles
    {
        public static void Run()
        {
            //ExStart:ExtractContentBetweenParagraphStyles
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

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
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            dstDoc.Save(dataDir);
            //ExEnd:ExtractContentBetweenParagraphStyles
            Console.WriteLine("\nExtracted content betweenn the paragraph styles successfully.\nFile saved at " + dataDir);
        }
    }
}
