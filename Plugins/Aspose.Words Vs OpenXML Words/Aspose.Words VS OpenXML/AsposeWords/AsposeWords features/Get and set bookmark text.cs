// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class GetAndSetBookmarkText : TestUtil
    {
        [Test]
        public void GetAndSetBookmarkTextFeature()
        {
            Document doc = new Document(MyDir + "Get and set bookmark text.docx");

            // Rename a bookmark and edit its text.
            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];
            bookmark.Name = "RenamedBookmark";
            bookmark.Text = "This is a new bookmarked text.";

            doc.Save(ArtifactsDir + "Get and set bookmark text - Aspose.Words.docx");
        }
    }
}
