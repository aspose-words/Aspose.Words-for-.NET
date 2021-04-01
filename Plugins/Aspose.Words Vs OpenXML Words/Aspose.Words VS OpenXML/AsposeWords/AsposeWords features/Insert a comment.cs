// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class InsertAComment : TestUtil
    {
        [Test]
        public void InsertACommentFeature()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Comment comment = new Comment(doc);

            // Insert some text into the comment.
            Paragraph commentParagraph = new Paragraph(doc);
            commentParagraph.AppendChild(new Run(doc, "This is comment!!!"));
            comment.AppendChild(commentParagraph);

            // Create a "CommentRangeStart" and "CommentRangeEnd".
            int commentId = 0;
            CommentRangeStart start = new CommentRangeStart(doc, commentId);
            CommentRangeEnd end = new CommentRangeEnd(doc, commentId);

            builder.Write("This text is before the comment. ");

            // Insert comment and comment range start.
            builder.InsertNode(comment);
            builder.InsertNode(start);

            // Insert some more text.
            builder.Write("This text is commented. ");

            // Insert end of comment range.
            builder.InsertNode(end);

            builder.Write("This text is after the comment.");

            doc.Save(ArtifactsDir + "Insert a comment - Aspose.Words.docx");
        }
    }
}
