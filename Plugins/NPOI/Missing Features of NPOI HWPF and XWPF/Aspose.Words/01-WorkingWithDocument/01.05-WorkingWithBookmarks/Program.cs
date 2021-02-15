using System;
using System.IO;
using Aspose.Words;

namespace _01._05_WorkingWithBookmarks
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for an Aspose.Words license file in the local file system and apply it, if it exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                Aspose.Words.License license = new Aspose.Words.License();

                // Use the license from the bin/debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document();
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
