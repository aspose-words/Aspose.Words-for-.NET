using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Tables;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class BookmarkTable
    {
        public static void Run()
        {
            //ExStart:BookmarkTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();
            
            // Create empty document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // Insert a cell
            builder.InsertCell();

            // Start bookmark here after calling InsertCell
            builder.StartBookmark("MyBookmark");

            builder.Write("This is row 1 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.Write("This is row 1 cell 2");

            builder.EndRow();

            // Insert a cell
            builder.InsertCell();
            builder.Writeln("This is row 2 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.Writeln("This is row 2 cell 2");

            builder.EndRow();

            builder.EndTable();
            // End of bookmark
            builder.EndBookmark("MyBookmark");

            dataDir = dataDir + "Bookmark.Table_out_.doc";
            doc.Save(dataDir);
            //ExEnd:BookmarkTable
            Console.WriteLine("\nTable bookmarked successfully.\nFile saved at " + dataDir);
        }
        
    }
}
