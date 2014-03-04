// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.IO;
using System.Reflection;

using Aspose.Words;

namespace ProcessComments
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The sample infrastructure.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");

            //ExStart
            //ExId:ProcessComments_Main
            //ExSummary: The demo-code that illustrates the methods for the comments extraction and removal.
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

            // Save the document.
            doc.Save(dataDir + "Test File Out.doc");
            //ExEnd
        }

        /// <param name="doc">The source document.</param>
        //ExStart
        //ExFor:Comment.Author
        //ExFor:Comment.DateTime
        //ExId:ProcessComments_Extract_All
        //ExSummary:Extracts the author name, date&time and text of all comments in the document.
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
        //ExEnd

        /// <param name="doc">The source document.</param>
        /// <param name="authorName">The name of the comment's author.</param>
        //ExStart
        //ExId:ProcessComments_Extract_Author
        //ExSummary:Extracts the author name, date&time and text of the comments by the specified author.
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
        //ExEnd

        /// <param name="doc">The source document.</param>
        //ExStart
        //ExId:ProcessComments_Remove_All
        //ExSummary:Removes all comments in the document.
        static void RemoveComments(Document doc)
        {
            // Collect all comments in the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            // Remove all comments.
            comments.Clear();
        }
        //ExEnd

        /// <param name="doc">The source document.</param>
        /// <param name="authorName">The name of the comment's author.</param>
        //ExStart
        //ExId:ProcessComments_Remove_Author
        //ExSummary:Removes comments by the specified author.
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
        //ExEnd
    }
}