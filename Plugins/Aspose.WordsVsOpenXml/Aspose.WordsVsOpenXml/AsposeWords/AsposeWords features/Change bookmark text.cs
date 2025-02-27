// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class GetAndSetBookmarkText : TestUtil
    {
        [Test]
        public void BookmarkTextAsposeWords()
        {
            //ExStart:BookmarkTextAsposeWords
            //GistDesc:Get and set bookmark text using C#
            Document doc = new Document(MyDir + "Bookmark.docx");

            // Rename a bookmark and edit its text.
            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];
            bookmark.Text = "This is a new bookmarked text.";

            doc.Save(ArtifactsDir + "Bookmark text - Aspose.Words.docx");
            //ExEnd:BookmarkTextAsposeWords
        }
    }
}
