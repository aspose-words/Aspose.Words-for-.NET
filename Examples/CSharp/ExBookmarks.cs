//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using Aspose.Words;
using NUnit.Framework;

namespace Examples
{
    [TestFixture]
    public class ExBookmarks : ExBase
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
            Document doc = new Document(MyDir + "Bookmark.doc");

            // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

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
            Document doc = new Document(MyDir + "Bookmark.doc");

            // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

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
            Document doc = new Document(MyDir + "Bookmark.doc");
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
            Document doc = new Document(MyDir + "Bookmarks.doc");

            // By index.
            Bookmark bookmark1 = doc.Range.Bookmarks[0];
            
            // By name.
            Bookmark bookmark2 = doc.Range.Bookmarks["Bookmark2"];
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
            Document doc = new Document(MyDir + "Bookmarks.doc");
            // Remove a particular bookmark from the document.
            Bookmark bookmark = doc.Range.Bookmarks[0];
            doc.Range.Bookmarks.Remove(bookmark);

            // Remove a bookmark by specified name.
            doc.Range.Bookmarks.Remove("Bookmark2");

            // Remove a bookmark at the specified index.
            doc.Range.Bookmarks.RemoveAt(0);
            //ExEnd

            Assert.AreEqual(0, doc.Range.Bookmarks.Count);
        }

        [Test]
        public void BookmarksInsertBookmark()
        {
            //ExStart
            //ExId:BookmarksInsertBookmark
            //ExSummary:Shows how to create a new bookmark.
            Document doc = new Document();
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
            Document doc = new Document(MyDir + "Bookmark.doc");

            int count = doc.Range.Bookmarks.Count;
            //ExEnd

            Assert.AreEqual(1, count);
        }
    }
}
