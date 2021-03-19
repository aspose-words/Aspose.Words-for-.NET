using System;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            // Insert a comment.
            Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            comment.SetText("My comment.");

            // Place the comment at a node in the document's body.
            // This comment will show up at the location of its paragraph,
            // outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
            builder.CurrentParagraph.AppendChild(comment);

            // Collect all comments in the document.
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Remove all comments.
            comments.Clear();
        }
    }
}
