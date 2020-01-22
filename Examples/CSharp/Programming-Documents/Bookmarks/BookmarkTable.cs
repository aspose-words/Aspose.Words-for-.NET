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
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();

            InsertBookmarkTable(dataDir);
            BookmarkTableColumns(dataDir);
        }
        public static void InsertBookmarkTable(string dataDir)
        {
            // ExStart:BookmarkTable
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

            dataDir = dataDir + "Bookmark.Table_out.doc";
            doc.Save(dataDir);
            // ExEnd:BookmarkTable
            Console.WriteLine("\nTable bookmarked successfully.\nFile saved at " + dataDir);
        }

        public static void BookmarkTableColumns(string dataDir)
        {
            // ExStart:BookmarkTableColumns
            // Create empty document
            Document doc = new Document(dataDir + "Bookmark.Table_out.doc");
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                Console.WriteLine("Bookmark: {0}{1}", bookmark.Name, bookmark.IsColumn ? " (Column)" : "");
                if (bookmark.IsColumn)
                {
                    Row row = bookmark.BookmarkStart.GetAncestor(NodeType.Row) as Row;
                    if (row != null && bookmark.FirstColumn < row.Cells.Count)
                        Console.WriteLine(row.Cells[bookmark.FirstColumn].GetText().TrimEnd(ControlChar.CellChar));
                }
            }
            // ExEnd:BookmarkTableColumns
        }
    }
}
