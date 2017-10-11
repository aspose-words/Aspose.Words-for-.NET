using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CommentReply
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithComments();
            AddRemoveCommentReply(dataDir);
        }

        static void AddRemoveCommentReply(string dataDir)
        {
            // ExStart:AddRemoveCommentReply
            Document doc = new Document(dataDir + "TestFile.doc");
            Comment comment = (Comment)doc.GetChild(NodeType.Comment, 0, true);

            //Remove the reply
            comment.RemoveReply(comment.Replies[0]);

            //Add a reply to comment
            comment.AddReply("John Doe", "JD", new DateTime(2017, 9, 25, 12, 15, 0), "New reply");

            dataDir = dataDir + "TestFile_Out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:AddRemoveCommentReply   
            Console.WriteLine("\nComment's reply is removed successfully.\nFile saved at " + dataDir);
        }
    }
}
