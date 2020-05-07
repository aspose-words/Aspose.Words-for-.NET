// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Pdf.Facades;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

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
            //ExFor:BookmarksOutlineLevelCollection.Contains(System.String)
            //ExFor:BookmarksOutlineLevelCollection.Count
            //ExFor:BookmarksOutlineLevelCollection.IndexOfKey(System.String)
            //ExFor:BookmarksOutlineLevelCollection.Item(System.Int32)
            //ExFor:BookmarksOutlineLevelCollection.Item(System.String)
            //ExFor:BookmarksOutlineLevelCollection.Remove(System.String)
            //ExFor:BookmarksOutlineLevelCollection.RemoveAt(System.Int32)
            //ExFor:OutlineOptions.BookmarksOutlineLevels
            //ExSummary:Shows how to set outline levels for bookmarks.
            // Open a blank document, create a DocumentBuilder, and use the builder to add some text wrapped inside bookmarks
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Note that whitespaces in bookmark names will be converted into underscores when saved to Microsoft Word formats
            // such as .doc and .docx, but will be preserved in other formats like .pdf or .xps
            builder.StartBookmark("Bookmark 1");
            builder.Writeln("Text inside Bookmark 1.");

            builder.StartBookmark("Bookmark 2");
            builder.Writeln("Text inside Bookmark 1 and 2.");
            builder.EndBookmark("Bookmark 2");

            builder.Writeln("Text inside Bookmark 1.");
            builder.EndBookmark("Bookmark 1");

            builder.StartBookmark("Bookmark 3");
            builder.Writeln("Text inside Bookmark 3.");
            builder.EndBookmark("Bookmark 3");

            // We can specify outline levels for our bookmarks so that they show up in the table of contents and are indented by an amount
            // of space proportional to the indent level in a SaveOptions object
            // Some pdf/xps readers such as Google Chrome also allow the collapsing of all higher level bookmarks by adjacent lower level bookmarks
            // This feature applies to .pdf or .xps file formats, so only their respective SaveOptions subclasses will support it
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            BookmarksOutlineLevelCollection outlineLevels = pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels;

            outlineLevels.Add("Bookmark 1", 1);
            outlineLevels.Add("Bookmark 2", 2);
            outlineLevels.Add("Bookmark 3", 3);

            Assert.AreEqual(3, outlineLevels.Count);
            Assert.True(outlineLevels.Contains("Bookmark 1"));
            Assert.AreEqual(1, outlineLevels[0]);
            Assert.AreEqual(2, outlineLevels["Bookmark 2"]);
            Assert.AreEqual(2, outlineLevels.IndexOfKey("Bookmark 3"));

            // We can remove two elements so that only the outline level designation for "Bookmark 1" is left
            outlineLevels.RemoveAt(2);
            outlineLevels.Remove("Bookmark 2");

            // We have 9 bookmark levels to work with, and bookmark levels are also sorted in ascending order,
            // and get numbered in succession along that order
            // Practically this means that our three levels "1, 5, 9", will be seen as "1, 2, 3" in the output
            outlineLevels.Add("Bookmark 2", 5);
            outlineLevels.Add("Bookmark 3", 9);

            // Save the document as a .pdf and find links to the bookmarks and their outline levels
            doc.Save(ArtifactsDir + "BookmarksOutlineLevelCollection.BookmarkLevels.pdf", pdfSaveOptions);

            // We can empty this dictionary to remove the contents table
            outlineLevels.Clear();
            //ExEnd

            #if NET462 || NETCOREAPP2_1         
            // Bind pdf with Aspose.Pdf
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "BookmarksOutlineLevelCollection.BookmarkLevels.pdf");

            // Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(3, bookmarks.Count);

            // Assert that all the bookmarks title are with whitespaces
            Assert.AreEqual("Bookmark 1", bookmarks[0].Title);
            Assert.AreEqual("Bookmark 2", bookmarks[1].Title);
            Assert.AreEqual("Bookmark 3", bookmarks[2].Title);            
            #endif
        }
    }
}