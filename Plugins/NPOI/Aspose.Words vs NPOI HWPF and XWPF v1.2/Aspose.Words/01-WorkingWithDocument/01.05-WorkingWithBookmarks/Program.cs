using System;
using System.IO;
using Aspose.Words;

namespace _01._05_WorkingWithBookmarks
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document("../../data/document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ----- Set Bookmark

            builder.StartBookmark("AsposeBookmark");
            builder.Writeln("Text inside a bookmark.");
            builder.EndBookmark("AsposeBookmark");

            // ----- Get Bookmark
            
            // By index.
            Bookmark bookmark1 = doc.Range.Bookmarks[0];

            // By name.
            Bookmark bookmark2 = doc.Range.Bookmarks["AsposeBookmark"];

            doc.Save("AsposeBookmarks.doc");

        }
    }
}
