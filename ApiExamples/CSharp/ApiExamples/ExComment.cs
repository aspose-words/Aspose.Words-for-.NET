// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

            doc = DocumentHelper.SaveOpen(doc);
            Comment docComment = (Comment)doc.GetChild(NodeType.Comment, 0, true);

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
            //ExFor:Comment.Replies
            //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
            //ExSummary:Shows how to get all comments with all replies.
            Document doc = new Document(MyDir + "Comments.docx");

            // Get all comment from the document
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            Assert.AreEqual(12, comments.Count); //ExSkip

            // For all comments and replies we identify comment level and info about it
            foreach (Comment comment in comments.OfType<Comment>())
            {
                if (comment.Ancestor == null)
                {
                    Console.WriteLine("\nThis is a top-level comment");
                    Console.WriteLine("Comment author: " + comment.Author);
                    Console.WriteLine("Comment text: " + comment.GetText());

                    foreach (Comment commentReply in comment.Replies.OfType<Comment>())
                    {
                        Console.WriteLine("\n\tThis is a comment reply");
                        Console.WriteLine("\tReply author: " + commentReply.Author);
                        Console.WriteLine("\tReply text: " + commentReply.GetText());
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
            Document doc = new Document(MyDir + "Comments.docx");

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            Comment comment = (Comment)comments[0];
            Assert.AreEqual(2, comment.Replies.Count()); //ExSkip

            comment.RemoveAllReplies();
            Assert.AreEqual(0, comment.Replies.Count()); //ExSkip
            //ExEnd
        }

        [Test]
        public void RemoveCommentReply()
        {
            //ExStart
            //ExFor:Comment.RemoveReply(Comment)
            //ExFor:CommentCollection.Item(Int32)
            //ExSummary:Shows how to remove specific comment reply.
            Document doc = new Document(MyDir + "Comments.docx");

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            Comment parentComment = (Comment)comments[0];
            CommentCollection repliesCollection = parentComment.Replies;
            Assert.AreEqual(2, parentComment.Replies.Count()); //ExSkip

            // Remove the first reply to comment
            parentComment.RemoveReply(repliesCollection[0]);
            Assert.AreEqual(1, parentComment.Replies.Count()); //ExSkip
            //ExEnd
        }

        [Test]
        public void MarkCommentRepliesDone()
        {
            //ExStart
            //ExFor:Comment.Done
            //ExFor:CommentCollection
            //ExSummary:Shows how to mark comment as Done.
            Document doc = new Document(MyDir + "Comments.docx");

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            Comment comment = (Comment)comments[0];
            CommentCollection repliesCollection = comment.Replies;

            foreach (Comment childComment in repliesCollection)
            {
                if (!childComment.Done)
                {
                    // Update comment reply Done mark
                    childComment.Done = true;
                }
            }
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            comment = (Comment)doc.GetChildNodes(NodeType.Comment, true)[0];
            repliesCollection = comment.Replies;

            foreach (Comment childComment in repliesCollection)
            {
                Assert.True(childComment.Done);
            }
        }
        
        //ExStart
        //ExFor:Comment.Done
        //ExFor:Comment.#ctor(DocumentBase)
        //ExFor:Comment.Accept(DocumentVisitor)
        //ExFor:Comment.DateTime
        //ExFor:Comment.Id
        //ExFor:Comment.Initial
        //ExFor:CommentRangeEnd
        //ExFor:CommentRangeEnd.#ctor(DocumentBase,Int32)
        //ExFor:CommentRangeEnd.Accept(DocumentVisitor)
        //ExFor:CommentRangeEnd.Id
        //ExFor:CommentRangeStart
        //ExFor:CommentRangeStart.#ctor(DocumentBase,Int32)
        //ExFor:CommentRangeStart.Accept(DocumentVisitor)
        //ExFor:CommentRangeStart.Id
        //ExSummary:Shows how to create comments with replies and get all interested info.
        [Test] //ExSkip
        public void CreateCommentsAndPrintAllInfo()
        {
            Document doc = new Document();
            doc.RemoveAllChildren();

            Section sect = (Section)doc.AppendChild(new Section(doc));
            Body body = (Body)sect.AppendChild(new Body(doc));

            // Create a commented text with several comment replies
            for (int i = 0; i <= 10; i++)
            {
                Comment newComment = CreateComment(doc, "VDeryushev", "VD", DateTime.Now, "My test comment " + i);

                Paragraph para = (Paragraph)body.AppendChild(new Paragraph(doc));
                para.AppendChild(new CommentRangeStart(doc, newComment.Id));
                para.AppendChild(new Run(doc, "Commented text " + i));
                para.AppendChild(new CommentRangeEnd(doc, newComment.Id));
                para.AppendChild(newComment);
                
                for (int y = 0; y <= 2; y++)
                {
                    newComment.AddReply("John Doe", "JD", DateTime.Now, "New reply " + y);
                }
            }

            // Look at information of our comments
            PrintAllCommentInfo(ExtractComments(doc));
        }

        /// <summary>
        /// Create a new comment
        /// </summary>
        public static Comment CreateComment(Document doc, string author, string initials, DateTime dateTime, string text)
        {
            Comment newComment = new Comment(doc)
            {
                Author = author, Initial = initials, DateTime = dateTime
            };
            newComment.SetText(text);

            return newComment;
        }

        /// <summary>
        /// Extract comments from the document without replies.
        /// </summary>
        public static List<Comment> ExtractComments(Document doc)
        {
            List<Comment> collectedComments = new List<Comment>();
            
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                // All replies have ancestor, so we will add this check
                if (comment.Ancestor == null)
                {
                    collectedComments.Add(comment);
                }
            }

            return collectedComments;
        }

        /// <summary>
        /// Use an iterator and a visitor to print info of every comment from within a document.
        /// </summary>
        private static void PrintAllCommentInfo(List<Comment> comments)
        {
            // Create an object that inherits from the DocumentVisitor class
            CommentInfoPrinter commentVisitor = new CommentInfoPrinter();

            // Get the enumerator from the document's comment collection and iterate over the comments
            using (IEnumerator<Comment> enumerator = comments.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Comment currentComment = enumerator.Current;

                    // Accept our DocumentVisitor it to print information about our comments
                    if (currentComment != null)
                    {
                        // Get CommentRangeStart from our current comment and construct its information
                        CommentRangeStart commentRangeStart = (CommentRangeStart)currentComment.PreviousSibling.PreviousSibling.PreviousSibling;
                        commentRangeStart.Accept(commentVisitor);

                        // Construct current comment information
                        currentComment.Accept(commentVisitor);
                        
                        // Get CommentRangeEnd from our current comment and construct its information
                        CommentRangeEnd commentRangeEnd = (CommentRangeEnd)currentComment.PreviousSibling;
                        commentRangeEnd.Accept(commentVisitor);
                    }
                }

                // Output of all information received
                Console.WriteLine(commentVisitor.GetText());
            }
        }

        /// <summary>
        /// This Visitor implementation prints information and contents of all comments and comment ranges encountered in the document.
        /// </summary>
        public class CommentInfoPrinter : DocumentVisitor
        {
            public CommentInfoPrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideComment = false;
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideComment) IndentAndAppendLine("[Run] \"" + run.Text + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a CommentRangeStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentRangeStart)
            {
                IndentAndAppendLine("[Comment range start] ID: " + commentRangeStart.Id);
                mDocTraversalDepth++;
                mVisitorIsInsideComment = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a CommentRangeEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Comment range end] ID: " + commentRangeEnd.Id + "\n");
                mVisitorIsInsideComment = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Comment node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentStart(Comment comment)
            {
                IndentAndAppendLine(
                    $"[Comment start] For comment range ID {comment.Id}, By {comment.Author} on {comment.DateTime}");
                mDocTraversalDepth++;
                mVisitorIsInsideComment = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Comment node is ended in the document.
            /// </summary>
            public override VisitorAction VisitCommentEnd(Comment comment)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Comment end]");
                mVisitorIsInsideComment = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++)
                {
                    mBuilder.Append("|  ");
                }

                mBuilder.AppendLine(text);
            }

            private bool mVisitorIsInsideComment;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd
    }
}