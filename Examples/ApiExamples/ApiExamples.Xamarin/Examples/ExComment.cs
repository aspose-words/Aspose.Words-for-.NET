// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
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
            //ExSummary:Shows how to add a comment to a document, and then reply to it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            comment.SetText("My comment.");
            
            // Place the comment at a node in the document's body.
            // This comment will show up at the location of its paragraph,
            // outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
            builder.CurrentParagraph.AppendChild(comment);

            // Add a reply, which will show up under its parent comment.
            comment.AddReply("Joe Bloggs", "J.B.", DateTime.Now, "New reply");

            // Comments and replies are both Comment nodes.
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Comment, true).Count);

            // Comments that do not reply to other comments are "top-level". They have no ancestor comments.
            Assert.Null(comment.Ancestor);

            // Replies have an ancestor top-level comment.
            Assert.AreEqual(comment, comment.Replies[0].Ancestor);

            doc.Save(ArtifactsDir + "Comment.AddCommentWithReply.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Comment.AddCommentWithReply.docx");
            Comment docComment = (Comment)doc.GetChild(NodeType.Comment, 0, true);

            Assert.AreEqual(1, docComment.Count);
            Assert.AreEqual(1, comment.Replies.Count);

            Assert.AreEqual("\u0005My comment.\r", docComment.GetText());
            Assert.AreEqual("\u0005New reply\r", docComment.Replies[0].GetText());
        }

        [Test]
        public void PrintAllComments()
        {
            //ExStart
            //ExFor:Comment.Ancestor
            //ExFor:Comment.Author
            //ExFor:Comment.Replies
            //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
            //ExSummary:Shows how to print all of a document's comments and their replies.
            Document doc = new Document(MyDir + "Comments.docx");

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            Assert.AreEqual(12, comments.Count); //ExSkip

            // If a comment has no ancestor, it is a "top-level" comment as opposed to a reply-type comment.
            // Print all top-level comments along with any replies they may have.
            foreach (Comment comment in comments.OfType<Comment>().Where(c => c.Ancestor == null))
            {
                Console.WriteLine("Top-level comment:");
                Console.WriteLine($"\t\"{comment.GetText().Trim()}\", by {comment.Author}");
                Console.WriteLine($"Has {comment.Replies.Count} replies");
                foreach (Comment commentReply in comment.Replies)
                {
                    Console.WriteLine($"\t\"{commentReply.GetText().Trim()}\", by {commentReply.Author}");
                }
                Console.WriteLine();
            }
            //ExEnd
        }

        [Test]
        public void RemoveCommentReplies()
        {
            //ExStart
            //ExFor:Comment.RemoveAllReplies
            //ExFor:Comment.RemoveReply(Comment)
            //ExFor:CommentCollection.Item(Int32)
            //ExSummary:Shows how to remove comment replies.
            Document doc = new Document();

            Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            comment.SetText("My comment.");

            doc.FirstSection.Body.FirstParagraph.AppendChild(comment);
            
            comment.AddReply("Joe Bloggs", "J.B.", DateTime.Now, "New reply");
            comment.AddReply("Joe Bloggs", "J.B.", DateTime.Now, "Another reply");

            Assert.AreEqual(2, comment.Replies.Count()); 

            // Below are two ways of removing replies from a comment.
            // 1 -  Use the "RemoveReply" method to remove replies from a comment individually:
            comment.RemoveReply(comment.Replies[0]);

            Assert.AreEqual(1, comment.Replies.Count());

            // 2 -  Use the "RemoveAllReplies" method to remove all replies from a comment at once:
            comment.RemoveAllReplies();

            Assert.AreEqual(0, comment.Replies.Count()); 
            //ExEnd
        }

        [Test]
        public void Done()
        {
            //ExStart
            //ExFor:Comment.Done
            //ExFor:CommentCollection
            //ExSummary:Shows how to mark a comment as "done".
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Helo world!");

            // Insert a comment to point out an error. 
            Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            comment.SetText("Fix the spelling error!");
            doc.FirstSection.Body.FirstParagraph.AppendChild(comment);

            // Comments have a "Done" flag, which is set to "false" by default. 
            // If a comment suggests that we make a change within the document,
            // we can apply the change, and then also set the "Done" flag afterwards to indicate the correction.
            Assert.False(comment.Done);

            doc.FirstSection.Body.FirstParagraph.Runs[0].Text = "Hello world!";
            comment.Done = true;

            // Comments that are "done" will differentiate themselves
            // from ones that are not "done" with a faded text color.
            comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            comment.SetText("Add text to this paragraph.");
            builder.CurrentParagraph.AppendChild(comment);

            doc.Save(ArtifactsDir + "Comment.Done.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Comment.Done.docx");
            comment = (Comment)doc.GetChildNodes(NodeType.Comment, true)[0];

            Assert.True(comment.Done);
            Assert.AreEqual("\u0005Fix the spelling error!", comment.GetText().Trim());
            Assert.AreEqual("Hello world!", doc.FirstSection.Body.FirstParagraph.Runs[0].Text);
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
        //ExSummary:Shows how print the contents of all comments and their comment ranges using a document visitor.
        [Test] //ExSkip
        public void CreateCommentsAndPrintAllInfo()
        {
            Document doc = new Document();
            
            Comment newComment = new Comment(doc)
            {
                Author = "VDeryushev",
                Initial = "VD",
                DateTime = DateTime.Now
            };

            newComment.SetText("Comment regarding text.");

            // Add text to the document, warp it in a comment range, and then add your comment.
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.AppendChild(new CommentRangeStart(doc, newComment.Id));
            para.AppendChild(new Run(doc, "Commented text."));
            para.AppendChild(new CommentRangeEnd(doc, newComment.Id));
            para.AppendChild(newComment); 
            
            // Add two replies to the comment.
            newComment.AddReply("John Doe", "JD", DateTime.Now, "New reply.");
            newComment.AddReply("John Doe", "JD", DateTime.Now, "Another reply.");

            PrintAllCommentInfo(doc.GetChildNodes(NodeType.Comment, true));
        }
        
        /// <summary>
        /// Iterates over every top-level comment and prints its comment range, contents, and replies.
        /// </summary>
        private static void PrintAllCommentInfo(NodeCollection comments)
        {
            CommentInfoPrinter commentVisitor = new CommentInfoPrinter();

            // Iterate over all top-level comments. Unlike reply-type comments, top-level comments have no ancestor.
            foreach (Comment comment in comments.Where(c => ((Comment)c).Ancestor == null))
            {
                // First, visit the start of the comment range.
                CommentRangeStart commentRangeStart = (CommentRangeStart)comment.PreviousSibling.PreviousSibling.PreviousSibling;
                commentRangeStart.Accept(commentVisitor);

                // Then, visit the comment, and any replies that it may have.
                comment.Accept(commentVisitor);

                foreach (Comment reply in comment.Replies)
                    reply.Accept(commentVisitor);

                // Finally, visit the end of the comment range, and then print the visitor's text contents.
                CommentRangeEnd commentRangeEnd = (CommentRangeEnd)comment.PreviousSibling;
                commentRangeEnd.Accept(commentVisitor);

                Console.WriteLine(commentVisitor.GetText());
            }
        }

        /// <summary>
        /// Prints information and contents of all comments and comment ranges encountered in the document.
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