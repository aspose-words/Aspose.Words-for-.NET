using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenBlockLevelNodes
    {
        public static void Run()
        {
            //ExStart:ExtractContentBetweenBlockLevelNodes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            Paragraph startPara = (Paragraph)doc.LastSection.GetChild(NodeType.Paragraph, 2, true);
            Table endTable = (Table)doc.LastSection.GetChild(NodeType.Table, 0, true);

            // Extract the content between these nodes in the document. Include these markers in the extraction.
            ArrayList extractedNodes = Common.ExtractContent(startPara, endTable, true);

            // Lets reverse the array to make inserting the content back into the document easier.
            extractedNodes.Reverse();

            while (extractedNodes.Count > 0)
            {
                // Insert the last node from the reversed list 
                endTable.ParentNode.InsertAfter((Node)extractedNodes[0], endTable);
                // Remove this node from the list after insertion.
                extractedNodes.RemoveAt(0);
            }
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the generated document to disk.
            doc.Save(dataDir);
            //ExEnd:ExtractContentBetweenBlockLevelNodes
            Console.WriteLine("\nExtracted content betweenn the block level nodes successfully.\nFile saved at " + dataDir);
        }
    }
}
