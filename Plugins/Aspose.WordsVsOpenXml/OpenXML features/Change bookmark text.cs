// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class ChangeBookmarkText : TestUtil
    {
        [Test]
        public void BookmarkTextOpenXml()
        {
            //ExStart:BookmarkTextOpenXml
            //GistId:a07e9ebecd60b1cbdfc1063ab58e87c6
            File.Copy(MyDir + "Bookmark.docx", ArtifactsDir + "Bookmark text - OpenXML.docx", true);

            // Open the original Wordprocessing document.
            using WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Bookmark text - OpenXML.docx", false);

            MainDocumentPart mainPart = doc.MainDocumentPart;

            // Find the bookmark in the new document
            BookmarkStart bookmark = mainPart.Document.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name == "MyBookmark");

            // Get the parent element of the bookmark
            OpenXmlElement parentElement = bookmark.Parent;

            // Create a new run with the new text
            Run newRun = new Run(new Text("This is a new bookmarked text."));

            // Replace the bookmark with the new text
            parentElement.RemoveAllChildren(); // Remove existing content
            parentElement.Append(newRun); // Add new text

            // Save changes to the new document
            mainPart.Document.Save();
            //ExEnd:BookmarkTextOpenXml
        }
    }
}
