// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class GetAndSetBookmarkText : TestUtil
    {
        [Test]
        public void BookmarkText()
        {
            // Open the original Wordprocessing document.
            using (WordprocessingDocument originalDocument = WordprocessingDocument.Open(MyDir + "Bookmark.docx", false))
            {
                // Create a new Wordprocessing document.
                using (WordprocessingDocument newDocument = WordprocessingDocument.Create(ArtifactsDir + "Bookmark text - OpenXML.docx", WordprocessingDocumentType.Document))
                {
                    // Add a main document part to the new document.
                    MainDocumentPart newMainPart = newDocument.AddMainDocumentPart();
                    newMainPart.Document = new Document(new Body());

                    // Copy content from the original document to the new document.
                    MainDocumentPart originalMainPart = originalDocument.MainDocumentPart;
                    newMainPart.Document.Body = (Body)originalMainPart.Document.Body.Clone();

                    // Find the bookmark in the new document
                    var bookmark = newMainPart.Document.Descendants<BookmarkStart>()
                        .FirstOrDefault(b => b.Name == "MyBookmark");

                    // Get the parent element of the bookmark
                    var parentElement = bookmark.Parent;

                    // Create a new run with the new text
                    Run newRun = new Run(new Text("This is a new bookmarked text."));

                    // Replace the bookmark with the new text
                    parentElement.RemoveAllChildren(); // Remove existing content
                    parentElement.Append(newRun); // Add new text

                    // Save changes to the new document
                    newMainPart.Document.Save();
                }
            }
        }
    }
}
