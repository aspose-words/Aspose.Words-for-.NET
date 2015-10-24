// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;

namespace Aspose_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create an empty document and DocumentBuilder object.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a Comment.
            Comment comment = new Comment(doc);
            // Insert some text into the comment.
            Paragraph commentParagraph = new Paragraph(doc);
            commentParagraph.AppendChild(new Run(doc, "This is comment!!!"));
            comment.AppendChild(commentParagraph);

            // Create CommentRangeStart and CommentRangeEnd.
            int commentId = 0;
            CommentRangeStart start = new CommentRangeStart(doc, commentId);
            CommentRangeEnd end = new CommentRangeEnd(doc, commentId);

            // Insert some text into the document.
            builder.Write("This is text before comment ");
            // Insert comment and comment range start.
            builder.InsertNode(comment);
            builder.InsertNode(start);
            // Insert some more text.
            builder.Write("This is commented text ");
            // Insert end of comment range.
            builder.InsertNode(end);
            // And finaly insert some more text.
            builder.Write("This is text aftr comment");

            // Save output document.
            doc.Save("Insert a Comment in Word Processing document.docx");
        }
    }
}
