using System;
using System.Collections;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Comments
{
    class ProcessComments
    {
        public static void Run()
        {
            //ExStart:ProcessComments
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithComments();
            string fileName = "TestFile.doc";

            // Open the document.
            Document doc = new Document(dataDir + fileName);

            // Extract the information about the comments of all the authors.
            foreach (string comment in ExtractComments(doc))
                Console.Write(comment);

            // Remove comments by the "pm" author.
            RemoveComments(doc, "pm");
            Console.WriteLine("Comments from \"pm\" are removed!");

            // Extract the information about the comments of the "ks" author.
            foreach (string comment in ExtractComments(doc, "ks"))
                Console.Write(comment);

            // Remove all comments.
            RemoveComments(doc);
            Console.WriteLine("All comments are removed!");

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document.
            doc.Save(dataDir);
            //ExEnd:ProcessComments
            Console.WriteLine("\nComments extracted and removed successfully.\nFile saved at " + dataDir);
        }
        //ExStart:ExtractComments
        static ArrayList ExtractComments(Document doc)
        {
            ArrayList collectedComments = new ArrayList();
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            // Look through all comments and gather information about them.
            foreach (Comment comment in comments)
            {
                collectedComments.Add(comment.Author + " " + comment.DateTime + " " + comment.ToString(SaveFormat.Text));
            }
            return collectedComments;
        }
        //ExEnd:ExtractComments
        //ExStart:ExtractCommentsByAuthor
        static ArrayList ExtractComments(Document doc, string authorName)
        {
            ArrayList collectedComments = new ArrayList();
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            // Look through all comments and gather information about those written by the authorName author.
            foreach (Comment comment in comments)
            {
                if (comment.Author == authorName)
                    collectedComments.Add(comment.Author + " " + comment.DateTime + " " + comment.ToString(SaveFormat.Text));
            }
            return collectedComments;
        }
        //ExEnd:ExtractCommentsByAuthor
        //ExStart:RemoveComments
        static void RemoveComments(Document doc)
        {
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            // Remove all comments.
            comments.Clear();
        }
        //ExEnd:RemoveComments
        //ExStart:RemoveCommentsByAuthor
        static void RemoveComments(Document doc, string authorName)
        {
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            // Look through all comments and remove those written by the authorName author.
            for (int i = comments.Count - 1; i >= 0; i--)
            {
                Comment comment = (Comment)comments[i];
                if (comment.Author == authorName)
                    comment.Remove();
            }
        }
        //ExEnd:RemoveCommentsByAuthor
    }
}
