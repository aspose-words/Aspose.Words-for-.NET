// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using NUnit.Framework;
using System.IO;
using Aspose.Words;
using Aspose.Pdf.Facades;
using Aspose.Words.Saving;

namespace ApiExamples
{
    [TestFixture]
    public class ExBookmarks : ApiExampleBase
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
            Document doc = new Document(MyDir + "Bookmark.doc");

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
            Document doc = new Document(MyDir + "Bookmarks.doc");

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

        [Test]
        public void CreateBookmarkWithNodes()
        {
            //ExStart
            //ExFor:BookmarkStart
            //ExFor:BookmarkStart.#ctor
            //ExFor:BookmarkEnd
            //ExFor:BookmarkEnd.#ctor
            //ExSummary:Shows how to create a bookmark by inserting bookmark start and end nodes.
            Document doc = new Document();

            // An empty document has just one empty paragraph by default.
            Paragraph p = doc.FirstSection.Body.FirstParagraph;

            p.AppendChild(new Run(doc, "Text before bookmark. "));

            p.AppendChild(new BookmarkStart(doc, "My bookmark"));
            p.AppendChild(new Run(doc, "Text inside bookmark. "));
            p.AppendChild(new BookmarkEnd(doc, "My bookmark"));

            p.AppendChild(new Run(doc, "Text after bookmark."));

            doc.Save(MyDir + @"\Artifacts\Bookmarks.CreateBookmarkWithNodes.doc");
            //ExEnd

            Assert.AreEqual(doc.Range.Bookmarks["My bookmark"].Text, "Text inside bookmark. ");
        }

        [Test]
        public void ReplaceBookmarkUnderscoresWithWhitespaces()
        {
            //ExStart
            //ExFor:Bookmark.Name
            //ExSummary:Shows how to replace elements in bookmark name
            Document doc = new Document(MyDir + "Bookmarks.Replace.docx");

            Assert.AreEqual("My_Bookmark", doc.Range.Bookmarks[0].Name); //ExSkip

            //MS Word document does not support bookmark names with whitespaces by default. 
            //If you have document which contains bookmark names with underscores, you can simply replace them to whitespaces.
            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
            {
                bookmark.Name = bookmark.Name.Replace("_", " ");
            }
            //ExEnd

            Assert.AreEqual("My Bookmark", doc.Range.Bookmarks[0].Name); //Check that our replace was correct
        }

        [Test]
        public void AllowToAddBookmarksWithWhiteSpaces()
        {
            //ExStart
            //ExFor:OutlineOptions.BookmarksOutlineLevels
            //ExFor:BookmarksOutlineLevelCollection.Add(String, Int32)
            //ExSummary:Shows how adding bookmarks outlines with whitespaces(pdf, xps)
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Add bookmarks with whitespaces. MS Word formats (like doc, docx) does not support bookmarks with whitespaces by default 
            //and all whitespaces in the bookmarks were replaced with underscores. If you need to use bookmarks in PDF or XPS outlines, you can use them with whitespaces.
            builder.StartBookmark("My Bookmark");
            builder.Writeln("Text inside a bookmark.");

            builder.StartBookmark("Nested Bookmark");
            builder.Writeln("Text inside a NestedBookmark.");
            builder.EndBookmark("Nested Bookmark");

            builder.Writeln("Text after Nested Bookmark.");
            builder.EndBookmark("My Bookmark");

            //Specify bookmarks outline level. If you are using xps format, just use XpsSaveOptions.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1);
            pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2);

            doc.Save(MyDir + @"\Artifacts\Bookmarks.WhiteSpaces Out.pdf", pdfSaveOptions);
            //ExEnd

            //Bind pdf with Aspose.Pdf
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(MyDir + @"\Artifacts\Bookmarks.WhiteSpaces Out.pdf");

            //Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(2, bookmarks.Count);

            //Assert that all the bookmarks title are with whitespaces
            Assert.AreEqual("My Bookmark", bookmarks[0].Title);
            Assert.AreEqual("Nested Bookmark", bookmarks[1].Title);
        }
    }
}