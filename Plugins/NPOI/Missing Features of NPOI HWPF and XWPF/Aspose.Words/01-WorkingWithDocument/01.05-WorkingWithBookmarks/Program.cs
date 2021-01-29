using Aspose.Words;

namespace _01._05_WorkingWithBookmarks
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use the document builder to insert a bookmark which encases text.
            builder.StartBookmark("AsposeBookmark");
            builder.Writeln("Text inside a bookmark.");
            builder.EndBookmark("AsposeBookmark");

            // Below are two ways of accessing a bookmark in a document.
            // 1 -  By index:
            Bookmark bookmark1 = doc.Range.Bookmarks[0];

            // 2 -  By name:
            Bookmark bookmark2 = doc.Range.Bookmarks["AsposeBookmark"];

            doc.Save("WorkingWithBookmarks.docx");
        }
    }
}
