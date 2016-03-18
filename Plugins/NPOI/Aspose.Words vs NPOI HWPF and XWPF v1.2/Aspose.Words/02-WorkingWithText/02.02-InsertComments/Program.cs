using System;
using System.Collections.Generic;
using System.Text; using Aspose.Words;

namespace _02._02_InsertComments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Some text is added.");

            Comment comment = new Comment(doc, "Aspose", "As", new DateTime());
            builder.CurrentParagraph.AppendChild(comment);
            comment.Paragraphs.Add(new Paragraph(doc));
            comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));

            doc.Save("insertedComments.doc");
        }
    }
}
