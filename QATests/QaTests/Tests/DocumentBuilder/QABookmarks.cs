using System.IO;
using Aspose.Pdf.Facades;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace QaTests.Tests
{
    [TestFixture]
    internal class QaBookmarks : QaTestsBase
    {
        [Test]
        public void BookmarkWhiteSpacesPdf()
        {
            Document doc = new Document();
            
            InsertBookmarks(doc);

            //Save document with pdf save options
            doc.Save(MyDir + "Bookmark_WhiteSpaces_OUT.pdf", AddBookmarkSaveOptions(SaveFormat.Pdf));

            //Bind pdf with Aspose PDF
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(MyDir + "Bookmark_WhiteSpaces_OUT.pdf");

            //Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(3, bookmarks.Count);

            //Assert that all the bookmarks title are with witespaces
            Assert.AreEqual("My Bookmark", bookmarks[0].Title);
            Assert.AreEqual("Nested Bookmark", bookmarks[1].Title);

            //Assert that the bookmark title without witespaces
            Assert.AreEqual("Bookmark_WithoutWhiteSpaces", bookmarks[2].Title);
        }

        [Test]
        public void BookmarkWhiteSpacesXps()
        {
            Document doc = new Document();

            InsertBookmarks(doc);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, AddBookmarkSaveOptions(SaveFormat.Xps));

            //Get bookmarks from the document
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            Assert.AreEqual(3, bookmarks.Count);

            //Assert that all the bookmarks title are with witespaces
            Assert.AreEqual("My Bookmark", bookmarks[0].Name);
            Assert.AreEqual("Nested Bookmark", bookmarks[1].Name);

            //Assert that the bookmark title without witespaces
            Assert.AreEqual("Bookmark_WithoutWhiteSpaces", bookmarks[2].Name);
        }

        [Test]
        public void BookmarkWhiteSpacesSwf()
        {
            Document doc = new Document();

            InsertBookmarks(doc);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, AddBookmarkSaveOptions(SaveFormat.Swf));

            //Get bookmarks from the document
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            Assert.AreEqual(3, bookmarks.Count);

            //Assert that all the bookmarks title are with witespaces
            Assert.AreEqual("My Bookmark", bookmarks[0].Name);
            Assert.AreEqual("Nested Bookmark", bookmarks[1].Name);

            //Assert that the bookmark title without witespaces
            Assert.AreEqual("Bookmark_WithoutWhiteSpaces", bookmarks[2].Name);
        }

        private static void InsertBookmarks(Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("My Bookmark");
            builder.Writeln("Text inside a bookmark.");

            builder.StartBookmark("Nested Bookmark");
            builder.Writeln("Text inside a NestedBookmark.");
            builder.EndBookmark("Nested Bookmark");

            builder.Writeln("Text after Nested Bookmark.");
            builder.EndBookmark("My Bookmark");

            builder.StartBookmark("Bookmark_WithoutWhiteSpaces");
            builder.Writeln("Text inside a NestedBookmark.");
            builder.EndBookmark("Bookmark_WithoutWhiteSpaces");
        }

        private static SaveOptions AddBookmarkSaveOptions(SaveFormat saveFormat)
        {
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            XpsSaveOptions xpsSaveOptions = new XpsSaveOptions();
            SwfSaveOptions swfSaveOptions = new SwfSaveOptions();

            switch (saveFormat)
            {
                case SaveFormat.Pdf:

                    //Add bookmarks to the document
                    pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1);
                    pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2);
                    pdfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Bookmark_WithoutWhiteSpaces", 3);

                    return pdfSaveOptions;

                case SaveFormat.Xps:

                    //Add bookmarks to the document
                    xpsSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1);
                    xpsSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2);
                    xpsSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Bookmark_WithoutWhiteSpaces", 3);

                    return xpsSaveOptions;

                case SaveFormat.Swf:

                    //Add bookmarks to the document
                    swfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1);
                    swfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2);
                    swfSaveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Bookmark_WithoutWhiteSpaces", 3);

                    return swfSaveOptions;
            }

            return null;
        }
    }
}
