using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Saving;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class CreateBookmark
    {
        public static void Run()
        {
            //ExStart:CreateBookmark
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("My Bookmark");
            builder.Writeln("Text inside a bookmark.");

            builder.StartBookmark("Nested Bookmark");
            builder.Writeln("Text inside a NestedBookmark.");
            builder.EndBookmark("Nested Bookmark");

            builder.Writeln("Text after Nested Bookmark.");
            builder.EndBookmark("My Bookmark");


            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1);
            options.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2);

            dataDir = dataDir + "Create.Bookmark_out_.pdf";
            doc.Save(dataDir, options);
            //ExEnd:CreateBookmark
            Console.WriteLine("\nBookmark created successfully.\nFile saved at " + dataDir);
        }
        
    }
}
