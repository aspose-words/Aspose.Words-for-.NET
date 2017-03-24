using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Comments
{
    class RemoveRegionText
    {
        public static void Run()
        {
            // ExStart:RemoveRegionText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithComments();
            string fileName = "TestFile.doc";
            
            // Open the document.
            Document doc = new Document(dataDir + fileName);

            CommentRangeStart commentStart = (CommentRangeStart)doc.GetChild(NodeType.CommentRangeStart, 0, true);
            CommentRangeEnd commentEnd = (CommentRangeEnd)doc.GetChild(NodeType.CommentRangeEnd, 0, true);

            Node currentNode = commentStart;
            Boolean isRemoving = true;
            while (currentNode != null && isRemoving)
            {
                if (currentNode.NodeType == NodeType.CommentRangeEnd)
                    isRemoving = false;

                Node nextNode = currentNode.NextPreOrder(doc);
                currentNode.Remove();
                currentNode = nextNode;
            }
            
            dataDir = dataDir + "RemoveRegionText_out.doc";
            // Save the document.
            doc.Save(dataDir);
            // ExEnd:RemoveRegionText
            Console.WriteLine("\nComments added successfully.\nFile saved at " + dataDir);
        }
    }
}
