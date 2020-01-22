// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using NUnit.Framework;
using Aspose.Words;
using Aspose.Words.Tables;
using Bookmark = Aspose.Words.Bookmark;

namespace ApiExamples
{
    [TestFixture]
    public class ExBookmarks : ApiExampleBase
    {
        //ExStart
        //ExFor:Bookmark
        //ExFor:Bookmark.Name
        //ExFor:Bookmark.Text
        //ExFor:Bookmark.Remove
        //ExFor:Bookmark.BookmarkStart
        //ExFor:Bookmark.BookmarkEnd
        //ExFor:BookmarkStart
        //ExFor:BookmarkStart.#ctor
        //ExFor:BookmarkEnd
        //ExFor:BookmarkEnd.#ctor
        //ExFor:BookmarkStart.Accept(DocumentVisitor)
        //ExFor:BookmarkEnd.Accept(DocumentVisitor)
        //ExFor:BookmarkStart.Bookmark
        //ExFor:BookmarkStart.GetText
        //ExFor:BookmarkStart.Name
        //ExFor:BookmarkEnd.Name
        //ExFor:BookmarkCollection
        //ExFor:BookmarkCollection.Item(Int32)
        //ExFor:BookmarkCollection.Item(String)
        //ExFor:BookmarkCollection.Count
        //ExFor:BookmarkCollection.GetEnumerator
        //ExFor:Range.Bookmarks
        //ExFor:DocumentVisitor.VisitBookmarkStart 
        //ExFor:DocumentVisitor.VisitBookmarkEnd
        //ExSummary:Shows how to add bookmarks and update their contents.
        [Test] //ExSkip
        public void CreateUpdateAndPrintBookmarks()
        {
            // Create a document with 3 bookmarks: "MyBookmark 1", "MyBookmark 2", "MyBookmark 3"
            Document doc = CreateDocumentWithBookmarks();
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            // Check that we have 3 bookmarks
            Assert.AreEqual(3, bookmarks.Count);
            Assert.AreEqual("MyBookmark 1", bookmarks[0].Name); //ExSkip
            Assert.AreEqual("Text content of MyBookmark 2", bookmarks[1].Text); //ExSkip

            // Look at initial values of our bookmarks
            PrintAllBookmarkInfo(bookmarks);

            // Obtain bookmarks from a bookmark collection by index/name and update their values
            bookmarks[0].Name = "Updated name of " + bookmarks[0].Name;
            bookmarks["MyBookmark 2"].Text = "Updated text content of " + bookmarks[1].Name;
            // Remove the latest bookmark
            // The bookmarked text is not deleted
            bookmarks[2].Remove();

            bookmarks = doc.Range.Bookmarks;
            // Check that we have 2 bookmarks after the latest bookmark was deleted
            Assert.AreEqual(2, bookmarks.Count);
            Assert.AreEqual("Updated name of MyBookmark 1", bookmarks[0].Name); //ExSkip
            Assert.AreEqual("Updated text content of MyBookmark 2", bookmarks[1].Text); //ExSkip

            // Look at updated values of our bookmarks
            PrintAllBookmarkInfo(bookmarks);
        }

        /// <summary>
        /// Create a document with bookmarks using the start and end nodes.
        /// </summary>
        private static Document CreateDocumentWithBookmarks()
        {
            DocumentBuilder builder = new DocumentBuilder();
            Document doc = builder.Document;

            // An empty document has just one empty paragraph by default
            Paragraph p = doc.FirstSection.Body.FirstParagraph;

            // Add several bookmarks to the document
            for (int i = 1; i <= 3; i++)
            {
                string bookmarkName = "MyBookmark " + i;

                p.AppendChild(new Run(doc, "Text before bookmark."));

                p.AppendChild(new BookmarkStart(doc, bookmarkName));
                p.AppendChild(new Run(doc, "Text content of " + bookmarkName));
                p.AppendChild(new BookmarkEnd(doc, bookmarkName));

                p.AppendChild(new Run(doc, "Text after bookmark.\r\n"));
            }

            return builder.Document;
        }

        /// <summary>
        /// Use an iterator and a visitor to print info of every bookmark from within a document.
        /// </summary>
        private static void PrintAllBookmarkInfo(BookmarkCollection bookmarks)
        {
            // Create a DocumentVisitor
            BookmarkInfoPrinter bookmarkVisitor = new BookmarkInfoPrinter();

            // Get the enumerator from the document's BookmarkCollection and iterate over the bookmarks
            using (IEnumerator<Bookmark> enumerator = bookmarks.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Bookmark currentBookmark = enumerator.Current;

                    // Accept our DocumentVisitor it to print information about our bookmarks
                    if (currentBookmark != null)
                    {
                        currentBookmark.BookmarkStart.Accept(bookmarkVisitor);
                        currentBookmark.BookmarkEnd.Accept(bookmarkVisitor);

                        // Prints a blank line
                        Console.WriteLine(currentBookmark.BookmarkStart.GetText());
                    }
                }
            }
        }

        /// <summary>
        /// Visitor that prints bookmark information to the console.
        /// </summary>
        public class BookmarkInfoPrinter : DocumentVisitor
        {
            public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
            {
                Console.WriteLine("BookmarkStart name: \"{0}\", Content: \"{1}\"", bookmarkStart.Name,
                    bookmarkStart.Bookmark.Text);
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
            {
                Console.WriteLine("BookmarkEnd name: \"{0}\"", bookmarkEnd.Name);
                return VisitorAction.Continue;
            }
        }
        //ExEnd

        [Test]
        public void TableColumnBookmarks()
        {
            //ExStart
            //ExFor:Bookmark.IsColumn
            //ExFor:Bookmark.FirstColumn
            //ExFor:Bookmark.LastColumn
            //ExSummary:Shows how to get information about table column bookmark.
            Document doc = new Document(MyDir + "TableColumnBookmark.doc");
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                Console.WriteLine("Bookmark: {0}{1}", bookmark.Name, bookmark.IsColumn ? " (Column)" : "");
                if (bookmark.IsColumn)
                {
                    if (bookmark.BookmarkStart.GetAncestor(NodeType.Row) is Row row &&
                        bookmark.FirstColumn < row.Cells.Count)
                    {
                        // Print text from the first and last cells containing in bookmark
                        Console.WriteLine(row.Cells[bookmark.FirstColumn].GetText().TrimEnd(ControlChar.CellChar));
                        Console.WriteLine(row.Cells[bookmark.LastColumn].GetText().TrimEnd(ControlChar.CellChar));
                    }
                }
            }
            //ExEnd

            Bookmark firstTableColumnBookmark = doc.Range.Bookmarks["FirstTableColumnBookmark"];
            Bookmark secondTableColumnBookmark = doc.Range.Bookmarks["SecondTableColumnBookmark"];

            Assert.IsTrue(firstTableColumnBookmark.IsColumn);
            Assert.AreEqual(1, firstTableColumnBookmark.FirstColumn);
            Assert.AreEqual(3, firstTableColumnBookmark.LastColumn);

            Assert.IsTrue(secondTableColumnBookmark.IsColumn);
            Assert.AreEqual(0, secondTableColumnBookmark.FirstColumn);
            Assert.AreEqual(3, secondTableColumnBookmark.LastColumn);
        }

        [Test]
        public void ClearBookmarks()
        {
            //ExStart
            //ExFor:BookmarkCollection.Clear
            //ExSummary:Shows how to remove all bookmarks from a document.
            // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
            Document doc = new Document(MyDir + "Bookmarks.docx");

            // Remove all bookmarks from the document
            // The bookmarked text is not deleted
            doc.Range.Bookmarks.Clear();
            //ExEnd

            // Verify that the bookmarks were removed
            Assert.AreEqual(0, doc.Range.Bookmarks.Count);
        }

        [Test]
        public void RemoveBookmarkFromBookmarkCollection()
        {
            //ExStart
            //ExFor:BookmarkCollection.Remove(Bookmark)
            //ExFor:BookmarkCollection.Remove(String)
            //ExFor:BookmarkCollection.RemoveAt
            //ExSummary:Shows how to remove bookmarks from a document using different methods.
            // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
            Document doc = new Document(MyDir + "Bookmarks.docx");

            // Remove a particular bookmark from the document
            Bookmark bookmark = doc.Range.Bookmarks[0];
            doc.Range.Bookmarks.Remove(bookmark);

            // Remove a bookmark by specified name
            doc.Range.Bookmarks.Remove("My_Bookmark2");

            // Remove a bookmark at the specified index
            doc.Range.Bookmarks.RemoveAt(0);
            //ExEnd

            // In docx we have additional hidden bookmark "_GoBack"
            // When we check bookmarks count, the result will be 1 instead of 0
            Assert.AreEqual(1, doc.Range.Bookmarks.Count);
        }

        [Test]
        public void ReplaceBookmarkUnderscoresWithWhitespaces()
        {
            //ExStart
            //ExFor:Bookmark.Name
            //ExSummary:Shows how to replace elements in bookmark name
            // Open a document with 3 bookmarks: "MyBookmark1", "My_Bookmark2", "MyBookmark3"
            Document doc = new Document(MyDir + "Bookmarks.docx");
            Assert.AreEqual("My_Bookmark2", doc.Range.Bookmarks[2].Name); //ExSkip

            // MS Word document does not support bookmark names with whitespaces by default
            // If you have document which contains bookmark names with underscores, you can simply replace them to whitespaces
            foreach (Bookmark bookmark in doc.Range.Bookmarks) bookmark.Name = bookmark.Name.Replace("_", " ");
            //ExEnd
        }
    }
}