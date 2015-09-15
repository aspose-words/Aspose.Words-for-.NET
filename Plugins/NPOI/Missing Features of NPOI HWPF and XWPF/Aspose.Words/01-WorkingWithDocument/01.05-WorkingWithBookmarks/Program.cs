using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;

namespace _01._05_WorkingWithBookmarks
{
    class Program
    {
        static void Main(string[] args)
        {
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
