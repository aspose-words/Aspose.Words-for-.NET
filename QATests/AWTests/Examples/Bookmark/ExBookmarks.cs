// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Bookmark
{
    [TestFixture]
    public class ExBookmarks : QaTestsBase
    {
        [Test]
        public void BookmarkNameAndText()
        {
            //ExStart
            //ExFor:Bookmark
            //ExFor:Bookmark.Name
            //ExFor:Bookmark.Text
            //ExFor:Range.Bookmarks
            //ExId:BookmarksGetNameSetText
            //ExSummary:Shows how to get or set bookmark name and text.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Bookmark.doc");

            // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
            Aspose.Words.Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

            // Get the name and text of the bookmark.
            string name = bookmark.Name;
            string text = bookmark.Text;

            // Set the name and text of the bookmark.
            bookmark.Name = "RenamedBookmark";
            bookmark.Text = "This is a new bookmarked text.";
            //ExEnd

            Assert.AreEqual("MyBookmark", name);
            Assert.AreEqual("This is a bookmarked text.", text);
        }

        [Test]
        public void BookmarkRemove()
        {
            //ExStart
            //ExFor:Bookmark.Remove
            //ExSummary:Shows how to remove a particular bookmark from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Bookmark.doc");

            // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
            Aspose.Words.Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

            // Remove the bookmark. The bookmarked text is not deleted.
            bookmark.Remove();
            //ExEnd

            // Verify that the bookmarks were removed from the document.
            Assert.AreEqual(0, doc.Range.Bookmarks.Count);
        }

        [Test]
        public void ClearBookmarks()
        {
            //ExStart
            //ExFor:BookmarkCollection.Clear
            //ExSummary:Shows how to remove all bookmarks from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Bookmark.doc");
            doc.Range.Bookmarks.Clear();
            //ExEnd

            // Verify that the bookmarks were removed
            Assert.AreEqual(0, doc.Range.Bookmarks.Count);
        }

        [Test]
        public void AccessBookmarks()
        {
            //ExStart
            //ExFor:BookmarkCollection
            //ExFor:BookmarkCollection.Item(Int32)
            //ExFor:BookmarkCollection.Item(String)
            //ExId:BookmarksAccess
            //ExSummary:Shows how to obtain bookmarks from a bookmark collection.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Bookmarks.doc");

            // By index.
            Aspose.Words.Bookmark bookmark1 = doc.Range.Bookmarks[0];
            
            // By name.
            Aspose.Words.Bookmark bookmark2 = doc.Range.Bookmarks["Bookmark2"];
            //ExEnd
        }

        [Test]
        public void BookmarkCollectionRemove()
        {
            //ExStart
            //ExFor:BookmarkCollection.Remove(Bookmark)
            //ExFor:BookmarkCollection.Remove(String)
            //ExFor:BookmarkCollection.RemoveAt
            //ExSummary:Demonstrates different methods of removing bookmarks from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Bookmarks.doc");
            // Remove a particular bookmark from the document.
            Aspose.Words.Bookmark bookmark = doc.Range.Bookmarks[0];
            doc.Range.Bookmarks.Remove(bookmark);

            // Remove a bookmark by specified name.
            doc.Range.Bookmarks.Remove("Bookmark2");

            // Remove a bookmark at the specified index.
            doc.Range.Bookmarks.RemoveAt(0);
            //ExEnd

            Assert.AreEqual(0, doc.Range.Bookmarks.Count);
        }

        [Test]
        public void BookmarksInsertBookmarkWithDocumentBuilder()
        {
            //ExStart
            //ExId:BookmarksInsertBookmark
            //ExSummary:Shows how to create a new bookmark.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("MyBookmark");
            builder.Writeln("Text inside a bookmark.");
            builder.EndBookmark("MyBookmark");
            //ExEnd
        }

        [Test]
        public void GetBookmarkCount()
        {
            //ExStart
            //ExFor:BookmarkCollection.Count
            //ExSummary:Shows how to count the number of bookmarks in a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Bookmark.doc");

            int count = doc.Range.Bookmarks.Count;
            //ExEnd

            Assert.AreEqual(1, count);
        }

        [Test]
        public void CreateBookmarkWithNodes()
        {
            //ExStart
            //ExFor:BookmarkStart
            //ExFor:BookmarkStart.#ctor
            //ExFor:BookmarkEnd
            //ExFor:BookmarkEnd.#ctor
            //ExSummary:Shows how to create a bookmark by inserting bookmark start and end nodes.
            Aspose.Words.Document doc = new Aspose.Words.Document();

            // An empty document has just one empty paragraph by default.
            Paragraph p = doc.FirstSection.Body.FirstParagraph;

            p.AppendChild(new Run(doc, "Text before bookmark. "));

            p.AppendChild(new BookmarkStart(doc, "My bookmark"));
            p.AppendChild(new Run(doc, "Text inside bookmark. "));
            p.AppendChild(new BookmarkEnd(doc, "My bookmark"));

            p.AppendChild(new Run(doc, "Text after bookmark."));

            doc.Save(ExDir + "Bookmarks.CreateBookmarkWithNodes.doc");

            Assert.AreEqual(doc.Range.Bookmarks["My bookmark"].Text, "Text inside bookmark. ");
            //ExEnd
        }
    }
}
