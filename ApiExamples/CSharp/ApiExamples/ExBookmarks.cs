// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Aspose.Words;
using Aspose.Words.Tables;
using Bookmark = Aspose.Words.Bookmark;

namespace ApiExamples
{
    [TestFixture]
    public class ExBookmarks : ApiExampleBase
    {
        [Test]
        public void Insert()
        {
            //ExStart
            //ExFor:Bookmark.Name
            //ExSummary:Shows how to insert a bookmark.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A valid bookmark has a name, a BookmarkStart, and a BookmarkEnd node.
            // Any whitespace in the names of bookmarks will be converted to underscores if we open the saved document with Microsoft Word. 
            // If we highlight the bookmark's name in Microsoft Word via Insert -> Links -> Bookmark, and press "Go To",
            // the cursor will jump to the text enclosed between the BookmarkStart and BookmarkEnd nodes.
            builder.StartBookmark("My Bookmark");
            builder.Write("Contents of MyBookmark.");
            builder.EndBookmark("My Bookmark");

            // Bookmarks are stored in this collection.
            Assert.AreEqual("My Bookmark", doc.Range.Bookmarks[0].Name);

            doc.Save(ArtifactsDir + "Bookmarks.Insert.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Bookmarks.Insert.docx");

            Assert.AreEqual("My Bookmark", doc.Range.Bookmarks[0].Name);
        }

        //ExStart
        //ExFor:Bookmark
        //ExFor:Bookmark.Name
        //ExFor:Bookmark.Text
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
        //ExFor:BookmarkCollection.GetEnumerator
        //ExFor:Range.Bookmarks
        //ExFor:DocumentVisitor.VisitBookmarkStart 
        //ExFor:DocumentVisitor.VisitBookmarkEnd
        //ExSummary:Shows how to add bookmarks and update their contents.
        [Test] //ExSkip
        public void CreateUpdateAndPrintBookmarks()
        {
            // Create a document with three bookmarks, then use a custom document visitor implementation to print their contents.
            Document doc = CreateDocumentWithBookmarks(3);
            BookmarkCollection bookmarks = doc.Range.Bookmarks;
            Assert.AreEqual(3, bookmarks.Count); //ExSkip

            PrintAllBookmarkInfo(bookmarks);
            
            // Bookmarks can be accessed in the bookmark collection by index or name, and their names can be updated.
            bookmarks[0].Name = $"{bookmarks[0].Name}_NewName";
            bookmarks["MyBookmark_2"].Text = $"Updated text contents of {bookmarks[1].Name}";

            // Print all bookmarks again to see updated values.
            PrintAllBookmarkInfo(bookmarks);
        }

        /// <summary>
        /// Create a document with a given number of bookmarks.
        /// </summary>
        private static Document CreateDocumentWithBookmarks(int numberOfBookmarks)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 1; i <= numberOfBookmarks; i++)
            {
                string bookmarkName = "MyBookmark_" + i;

                builder.Write("Text before bookmark.");
                builder.StartBookmark(bookmarkName);
                builder.Write($"Text inside {bookmarkName}.");
                builder.EndBookmark(bookmarkName);
                builder.Writeln("Text after bookmark.");
            }

            return doc;
        }

        /// <summary>
        /// Use an iterator and a visitor to print info of every bookmark in the collection.
        /// </summary>
        private static void PrintAllBookmarkInfo(BookmarkCollection bookmarks)
        {
            BookmarkInfoPrinter bookmarkVisitor = new BookmarkInfoPrinter();

            // Get each bookmark in the collection to accept a visitor that will print its contents.
            using (IEnumerator<Bookmark> enumerator = bookmarks.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Bookmark currentBookmark = enumerator.Current;

                    if (currentBookmark != null)
                    {
                        currentBookmark.BookmarkStart.Accept(bookmarkVisitor);
                        currentBookmark.BookmarkEnd.Accept(bookmarkVisitor);

                        Console.WriteLine(currentBookmark.BookmarkStart.GetText());
                    }
                }
            }
        }

        /// <summary>
        /// Prints contents of every visited bookmark to the console.
        /// </summary>
        public class BookmarkInfoPrinter : DocumentVisitor
        {
            public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
            {
                Console.WriteLine($"BookmarkStart name: \"{bookmarkStart.Name}\", Contents: \"{bookmarkStart.Bookmark.Text}\"");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
            {
                Console.WriteLine($"BookmarkEnd name: \"{bookmarkEnd.Name}\"");
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
            //ExSummary:Shows how to get information about table column bookmarks.
            Document doc = new Document(MyDir + "Table column bookmarks.doc");

            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                // If a bookmark encloses columns of a table, it is a table column bookmark, and its IsColumn flag set to true.
                Console.WriteLine($"Bookmark: {bookmark.Name}{(bookmark.IsColumn ? " (Column)" : "")}");
                if (bookmark.IsColumn)
                {
                    if (bookmark.BookmarkStart.GetAncestor(NodeType.Row) is Row row &&
                        bookmark.FirstColumn < row.Cells.Count)
                    {
                        // Print the contents of the first and last columns enclosed by the bookmark.
                        Console.WriteLine(row.Cells[bookmark.FirstColumn].GetText().TrimEnd(ControlChar.CellChar));
                        Console.WriteLine(row.Cells[bookmark.LastColumn].GetText().TrimEnd(ControlChar.CellChar));
                    }
                }
            }
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Bookmark firstTableColumnBookmark = doc.Range.Bookmarks["FirstTableColumnBookmark"];
            Bookmark secondTableColumnBookmark = doc.Range.Bookmarks["SecondTableColumnBookmark"];

            Assert.True(firstTableColumnBookmark.IsColumn);
            Assert.AreEqual(1, firstTableColumnBookmark.FirstColumn);
            Assert.AreEqual(3, firstTableColumnBookmark.LastColumn);

            Assert.True(secondTableColumnBookmark.IsColumn);
            Assert.AreEqual(0, secondTableColumnBookmark.FirstColumn);
            Assert.AreEqual(3, secondTableColumnBookmark.LastColumn);
        }

        [Test]
        public void Remove()
        {
            //ExStart
            //ExFor:BookmarkCollection.Clear
            //ExFor:BookmarkCollection.Count
            //ExFor:BookmarkCollection.Remove(Bookmark)
            //ExFor:BookmarkCollection.Remove(String)
            //ExFor:BookmarkCollection.RemoveAt
            //ExFor:Bookmark.Remove
            //ExSummary:Shows how to remove bookmarks from a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert five bookmarks with text inside their boundaries.
            for (int i = 1; i <= 5; i++)
            {
                string bookmarkName = "MyBookmark_" + i;

                builder.StartBookmark(bookmarkName);
                builder.Write($"Text inside {bookmarkName}.");
                builder.EndBookmark(bookmarkName);
                builder.InsertBreak(BreakType.ParagraphBreak);
            }

            // This collection stores bookmarks.
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            Assert.AreEqual(5, bookmarks.Count);

            // There are several ways of removing bookmarks.
            // 1 -  Calling the bookmark's Remove method:
            bookmarks["MyBookmark_1"].Remove();

            Assert.False(bookmarks.Any(b => b.Name == "MyBookmark_1"));

            // 2 -  Passing the bookmark to the collection's Remove method:
            Bookmark bookmark = doc.Range.Bookmarks[0];
            doc.Range.Bookmarks.Remove(bookmark);

            Assert.False(bookmarks.Any(b => b.Name == "MyBookmark_2"));
            
            // 3 -  Removing a bookmark from the collection by name:
            doc.Range.Bookmarks.Remove("MyBookmark_3");

            Assert.False(bookmarks.Any(b => b.Name == "MyBookmark_3"));

            // 4 -  Removing a bookmark at an index in the bookmark collection:
            doc.Range.Bookmarks.RemoveAt(0);

            Assert.False(bookmarks.Any(b => b.Name == "MyBookmark_4"));

            // We can clear the entire bookmark collection.
            bookmarks.Clear();

            // The text that was inside the bookmarks is still present in the document.
            Assert.That(bookmarks, Is.Empty);
            Assert.AreEqual("Text inside MyBookmark_1.\r" +
                            "Text inside MyBookmark_2.\r" +
                            "Text inside MyBookmark_3.\r" +
                            "Text inside MyBookmark_4.\r" +
                            "Text inside MyBookmark_5.", doc.GetText().Trim());
            //ExEnd
        }
    }
}