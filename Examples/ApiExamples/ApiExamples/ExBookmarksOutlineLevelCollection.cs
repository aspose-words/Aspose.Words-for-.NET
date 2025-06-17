// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;
using Aspose.Pdf.Facades;

namespace ApiExamples
{
    [TestFixture]
    public class ExBookmarksOutlineLevelCollection : ApiExampleBase
    {
        [Test]
        public void BookmarkLevels()
        {
            //ExStart
            //ExFor:BookmarksOutlineLevelCollection
            //ExFor:BookmarksOutlineLevelCollection.Add(String, Int32)
            //ExFor:BookmarksOutlineLevelCollection.Clear
            //ExFor:BookmarksOutlineLevelCollection.Contains(String)
            //ExFor:BookmarksOutlineLevelCollection.Count
            //ExFor:BookmarksOutlineLevelCollection.IndexOfKey(String)
            //ExFor:BookmarksOutlineLevelCollection.Item(Int32)
            //ExFor:BookmarksOutlineLevelCollection.Item(String)
            //ExFor:BookmarksOutlineLevelCollection.Remove(String)
            //ExFor:BookmarksOutlineLevelCollection.RemoveAt(Int32)
            //ExFor:OutlineOptions.BookmarksOutlineLevels
            //ExSummary:Shows how to set outline levels for bookmarks.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a bookmark with another bookmark nested inside it.
            builder.StartBookmark("Bookmark 1");
            builder.Writeln("Text inside Bookmark 1.");

            builder.StartBookmark("Bookmark 2");
            builder.Writeln("Text inside Bookmark 1 and 2.");
            builder.EndBookmark("Bookmark 2");

            builder.Writeln("Text inside Bookmark 1.");
            builder.EndBookmark("Bookmark 1");

            // Insert another bookmark.
            builder.StartBookmark("Bookmark 3");
            builder.Writeln("Text inside Bookmark 3.");
            builder.EndBookmark("Bookmark 3");

            // When saving to .pdf, bookmarks can be accessed via a drop-down menu and used as anchors by most readers.
            // Bookmarks can also have numeric values for outline levels,
            // enabling lower level outline entries to hide higher-level child entries when collapsed in the reader.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            BookmarksOutlineLevelCollection outlineLevels = pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels;

            outlineLevels.Add("Bookmark 1", 1);
            outlineLevels.Add("Bookmark 2", 2);
            outlineLevels.Add("Bookmark 3", 3);

            Assert.That(outlineLevels.Count, Is.EqualTo(3));
            Assert.That(outlineLevels.Contains("Bookmark 1"), Is.True);
            Assert.That(outlineLevels[0], Is.EqualTo(1));
            Assert.That(outlineLevels["Bookmark 2"], Is.EqualTo(2));
            Assert.That(outlineLevels.IndexOfKey("Bookmark 3"), Is.EqualTo(2));

            // We can remove two elements so that only the outline level designation for "Bookmark 1" is left.
            outlineLevels.RemoveAt(2);
            outlineLevels.Remove("Bookmark 2");

            // There are nine outline levels. Their numbering will be optimized during the save operation.
            // In this case, levels "5" and "9" will become "2" and "3".
            outlineLevels.Add("Bookmark 2", 5);
            outlineLevels.Add("Bookmark 3", 9);

            doc.Save(ArtifactsDir + "BookmarksOutlineLevelCollection.BookmarkLevels.pdf", pdfSaveOptions);

            // Emptying this collection will preserve the bookmarks and put them all on the same outline level.
            outlineLevels.Clear();
            //ExEnd
        }

        [Test]
        public void UsePdfBookmarkEditorForBookmarkLevels()
        {
            BookmarkLevels();

            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "BookmarksOutlineLevelCollection.BookmarkLevels.pdf");

            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.That(bookmarks.Count, Is.EqualTo(3));
            Assert.That(bookmarks[0].Title, Is.EqualTo("Bookmark 1"));
            Assert.That(bookmarks[1].Title, Is.EqualTo("Bookmark 2"));
            Assert.That(bookmarks[2].Title, Is.EqualTo("Bookmark 3"));
        }
    }
}
