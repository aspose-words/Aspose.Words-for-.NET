using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithComments : DocsExamplesBase
    {
        [Test]
        public void AddComments()
        {
            //ExStart:AddComments
            //ExStart:CreateSimpleDocumentUsingDocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Some text is added.");
            //ExEnd:CreateSimpleDocumentUsingDocumentBuilder
            
            Comment comment = new Comment(doc, "Awais Hafeez", "AH", DateTime.Today);

            builder.CurrentParagraph.AppendChild(comment);

            comment.Paragraphs.Add(new Paragraph(doc));
            comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));

            doc.Save(ArtifactsDir + "WorkingWithComments.AddComments.docx");
            //ExEnd:AddComments
        }

        [Test]
        public void AnchorComment()
        {
            //ExStart:AnchorComment
            Document doc = new Document();

            Paragraph para1 = new Paragraph(doc);
            Run run1 = new Run(doc, "Some ");
            Run run2 = new Run(doc, "text ");
            para1.AppendChild(run1);
            para1.AppendChild(run2);
            doc.FirstSection.Body.AppendChild(para1);

            Paragraph para2 = new Paragraph(doc);
            Run run3 = new Run(doc, "is ");
            Run run4 = new Run(doc, "added ");
            para2.AppendChild(run3);
            para2.AppendChild(run4);
            doc.FirstSection.Body.AppendChild(para2);

            Comment comment = new Comment(doc, "Awais Hafeez", "AH", DateTime.Today);
            comment.Paragraphs.Add(new Paragraph(doc));
            comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));

            CommentRangeStart commentRangeStart = new CommentRangeStart(doc, comment.Id);
            CommentRangeEnd commentRangeEnd = new CommentRangeEnd(doc, comment.Id);

            run1.ParentNode.InsertAfter(commentRangeStart, run1);
            run3.ParentNode.InsertAfter(commentRangeEnd, run3);
            commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd);

            doc.Save(ArtifactsDir + "WorkingWithComments.AnchorComment.doc");
            //ExEnd:AnchorComment
        }

        [Test]
        public void AddRemoveCommentReply()
        {
            //ExStart:AddRemoveCommentReply
            Document doc = new Document(MyDir + "Comments.docx");

            Comment comment = (Comment) doc.GetChild(NodeType.Comment, 0, true);
            comment.RemoveReply(comment.Replies[0]);

            comment.AddReply("John Doe", "JD", new DateTime(2017, 9, 25, 12, 15, 0), "New reply");

            doc.Save(ArtifactsDir + "WorkingWithComments.AddRemoveCommentReply.docx");
            //ExEnd:AddRemoveCommentReply
        }

        [Test]
        public void ProcessComments()
        {
            //ExStart:ProcessComments
            Document doc = new Document(MyDir + "Comments.docx");

            // Extract the information about the comments of all the authors.
            foreach (string comment in ExtractComments(doc))
                Console.Write(comment);

            // Remove comments by the "pm" author.
            RemoveComments(doc, "pm");
            Console.WriteLine("Comments from \"pm\" are removed!");

            // Extract the information about the comments of the "ks" author.
            foreach (string comment in ExtractComments(doc, "ks"))
                Console.Write(comment);

            // Read the comment's reply and resolve them.
            CommentResolvedAndReplies(doc);

            // Remove all comments.
            RemoveComments(doc);
            Console.WriteLine("All comments are removed!");

            doc.Save(ArtifactsDir + "WorkingWithComments.ProcessComments.docx");
            //ExEnd:ProcessComments
        }

        //ExStart:ExtractComments
        List<string> ExtractComments(Document doc)
        {
            List<string> collectedComments = new List<string>();
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                collectedComments.Add(comment.Author + " " + comment.DateTime + " " +
                                      comment.ToString(SaveFormat.Text));
            }

            return collectedComments;
        }
        //ExEnd:ExtractComments

        //ExStart:ExtractCommentsByAuthor
        List<string> ExtractComments(Document doc, string authorName)
        {
            List<string> collectedComments = new List<string>();
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                if (comment.Author == authorName)
                    collectedComments.Add(comment.Author + " " + comment.DateTime + " " +
                                          comment.ToString(SaveFormat.Text));
            }

            return collectedComments;
        }
        //ExEnd:ExtractCommentsByAuthor

        //ExStart:RemoveComments
        void RemoveComments(Document doc)
        {
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            comments.Clear();
        }
        //ExEnd:RemoveComments

        //ExStart:RemoveCommentsByAuthor
        void RemoveComments(Document doc, string authorName)
        {
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Look through all comments and remove those written by the authorName.
            for (int i = comments.Count - 1; i >= 0; i--)
            {
                Comment comment = (Comment) comments[i];
                if (comment.Author == authorName)
                    comment.Remove();
            }
        }
        //ExEnd:RemoveCommentsByAuthor

        //ExStart:CommentResolvedandReplies
        void CommentResolvedAndReplies(Document doc)
        {
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            Comment parentComment = (Comment) comments[0];
            foreach (Comment childComment in parentComment.Replies)
            {
                // Get comment parent and status.
                Console.WriteLine(childComment.Ancestor.Id);
                Console.WriteLine(childComment.Done);

                // And update comment Done mark.
                childComment.Done = true;
            }
        }
        //ExEnd:CommentResolvedandReplies
    }
}
