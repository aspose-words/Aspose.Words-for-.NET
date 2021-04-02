using Aspose.Words;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a blank document, and then a document builder which we will use to populate the document with content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A bookmark consists of a "BookmarkStart" node and a "BookmarkEnd" node
            // with matching names, as well as contents that these nodes enclose.
            builder.StartBookmark("MyBookmark");
            builder.Writeln("Hello world!");
            builder.EndBookmark("MyBookmark");

            // Save the document to the local file system.
            doc.Save("Adding Bookmark.docx");
        }
    }
}
