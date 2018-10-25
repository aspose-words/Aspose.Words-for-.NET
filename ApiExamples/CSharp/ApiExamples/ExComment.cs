// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExComment : ApiExampleBase
    {
        [Test]
        public void AddCommentWithReply()
        {
            //ExStart
            //ExFor:Comment
            //ExFor:Comment.SetText(String)
            //ExFor:Comment.Replies
            //ExFor:Comment.AddReply(String, String, DateTime, String)
            //ExSummary:Shows how to add a comment with a reply to a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create new comment
            Comment newComment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            newComment.SetText("My comment.");

            // Add this comment to a document node
            builder.CurrentParagraph.AppendChild(newComment);

            // Add comment reply
            newComment.AddReply("John Doe", "JD", new DateTime(2017, 9, 25, 12, 15, 0), "New reply");
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Comment docComment = (Comment) doc.GetChild(NodeType.Comment, 0, true);

            Assert.AreEqual(1, docComment.Count);
            Assert.AreEqual(1, newComment.Replies.Count);

            Assert.AreEqual("\u0005My comment.\r", docComment.GetText());
            Assert.AreEqual("\u0005New reply\r", docComment.Replies[0].GetText());
        }

        [Test]
        public void GetAllCommentsAndReplies()
        {
            //ExStart
            //ExFor:Comment.Ancestor
            //ExFor:Comment.Author
            //ExSummary:Shows how to get all comments with all replies.
            Document doc = new Document(MyDir + "Comment.Document.docx");

            // Get all comment from the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            Assert.AreEqual(12, comments.Count); //ExSkip

            // For all comments and replies we identify comment level and info about it
            foreach (Comment comment in comments.OfType<Comment>())
            {
                if (comment.Ancestor == null)
                {
                    Console.WriteLine("This is a top-level comment\n");

                    Console.WriteLine("Comment author: " + comment.Author);
                    Console.WriteLine("Comment text: " + comment.GetText());

                    foreach (Comment commentReply in comment.Replies.OfType<Comment>())
                    {
                        Console.WriteLine("This is a comment reply\n");

                        Console.WriteLine("Comment author: " + commentReply.Author);
                        Console.WriteLine("Comment text: " + commentReply.GetText());
                    }
                }
            }

            //ExEnd
        }

        [Test]
        public void RemoveCommentReplies()
        {
            //ExStart
            //ExFor:Comment.RemoveAllReplies
            //ExSummary:Shows how to remove comment replies.
            Document doc = new Document(MyDir + "Comment.Document.docx");

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            Comment comment = (Comment) comments[0];

            comment.RemoveAllReplies();
            //ExEnd
        }

        [Test]
        public void RemoveCommentReply()
        {
            //ExStart
            //ExFor:Comment.RemoveReply(Comment)
            //ExSummary:Shows how to remove specific comment reply.
            Document doc = new Document(MyDir + "Comment.Document.docx");

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            Comment parentComment = (Comment) comments[0];

            // Remove the first reply to comment
            parentComment.RemoveReply(parentComment.Replies[0]);
            //ExEnd
        }

        [Test]
        public void MarkCommentRepliesDone()
        {
            //ExStart
            //ExFor:Comment.Done
            //ExSummary:Shows how to mark comment as Done.
            Document doc = new Document(MyDir + "Comment.Document.docx");

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            Comment comment = (Comment) comments[0];

            foreach (Comment childComment in comment.Replies.OfType<Comment>())
            {
                if (!childComment.Done)
                {
                    // Update comment reply Done mark.
                    childComment.Done = true;
                }
            }

            //ExEnd
        }
    }
}