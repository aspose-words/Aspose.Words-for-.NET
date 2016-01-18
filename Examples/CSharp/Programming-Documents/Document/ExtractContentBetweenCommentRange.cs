using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenCommentRange
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            Document doc = new Document(dataDir + "TestFile.doc");

            // This is a quick way of getting both comment nodes.
            // Your code should have a proper method of retrieving each corresponding start and end node.
            CommentRangeStart commentStart = (CommentRangeStart)doc.GetChild(NodeType.CommentRangeStart, 0, true);
            CommentRangeEnd commentEnd = (CommentRangeEnd)doc.GetChild(NodeType.CommentRangeEnd, 0, true);

            // Firstly extract the content between these nodes including the comment as well. 
            ArrayList extractedNodesInclusive = Common.ExtractContent(commentStart, commentEnd, true);
            Document dstDoc = Common.GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(dataDir + "TestFile.CommentInclusive Out.doc");

            // Secondly extract the content between these nodes without the comment.
            ArrayList extractedNodesExclusive = Common.ExtractContent(commentStart, commentEnd, false);
            dstDoc = Common.GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(dataDir + "TestFile.CommentExclusive Out.doc");

            Console.WriteLine("\nExtracted content between the comment range successfully.\nFile saved at " + dataDir + "TestFile.CommentExclusive Out.doc");
        }
    }
}
