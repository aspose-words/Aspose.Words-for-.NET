//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;

using Aspose.Words;
using Aspose.Words.Tables;

namespace CSharp.Programming_Documents.Bookmarks
{
    class UntangleRowBookmarks
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();

            // Load a document.
            Document doc = new Document(dataDir + "TestDefect1352.doc");

            // This perform the custom task of putting the row bookmark ends into the same row with the bookmark starts.
            Untangle(doc);

            // Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
            DeleteRowByBookmark(doc, "ROW2");

            // This is just to check that the other bookmark was not damaged.
            if (doc.Range.Bookmarks["ROW1"].BookmarkEnd == null)
                throw new Exception("Wrong, the end of the bookmark was deleted.");

            // Save the finished document.
            doc.Save(dataDir + "TestDefect1352 Out.doc");

            Console.WriteLine("\nRow bookmark untangled successfully.\nFile saved at " + dataDir + "TestDefect1352 Out.doc");
        }

        private static void Untangle(Document doc)
        {
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                // Get the parent row of both the bookmark and bookmark end node.
                Row row1 = (Row)bookmark.BookmarkStart.GetAncestor(typeof(Row));
                Row row2 = (Row)bookmark.BookmarkEnd.GetAncestor(typeof(Row));

                // If both rows are found okay and the bookmark start and end are contained
                // in adjacent rows, then just move the bookmark end node to the end
                // of the last paragraph in the last cell of the top row.
                if ((row1 != null) && (row2 != null) && (row1.NextSibling == row2))
                    row1.LastCell.LastParagraph.AppendChild(bookmark.BookmarkEnd);
            }
        }

        private static void DeleteRowByBookmark(Document doc, string bookmarkName)
        {
            // Find the bookmark in the document. Exit if cannot find it.
            Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
            if (bookmark == null)
                return;

            // Get the parent row of the bookmark. Exit if the bookmark is not in a row.
            Row row = (Row)bookmark.BookmarkStart.GetAncestor(typeof(Row));
            if (row == null)
                return;

            // Remove the row.
            row.Remove();
        }
    }
}
