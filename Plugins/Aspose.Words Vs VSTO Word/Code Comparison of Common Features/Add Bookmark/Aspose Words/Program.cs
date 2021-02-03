using Aspose.Words;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string mypath = "";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("MyBookmark");
            builder.Writeln("Text inside a bookmark.");
            builder.EndBookmark("MyBookmark");
            doc.Save(mypath + "Adding Bookmark.doc");
        }
    }
}
