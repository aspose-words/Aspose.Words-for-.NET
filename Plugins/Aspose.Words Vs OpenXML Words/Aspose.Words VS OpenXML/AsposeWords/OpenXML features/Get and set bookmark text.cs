// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class GetAndSetBookmarkText : TestUtil
    {
        [Test]
        public void GetAndSetBookmarkTextFeature()
        {
            IDictionary<string, BookmarkStart> bookmarkMap = new Dictionary<string, BookmarkStart>();

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(MyDir + "Get and set bookmark text.docx", true))
            {
                foreach (BookmarkStart bookmarkStart in wordDocument.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                {
                    bookmarkMap[bookmarkStart.Name] = bookmarkStart;

                    foreach (BookmarkStart bookmark in bookmarkMap.Values)
                    {
                        Run bookmarkText = bookmark.NextSibling<Run>();

                        if (bookmarkText != null)
                            bookmarkText.GetFirstChild<Text>().Text = "Test";
                    }
                }
            }
        }
    }
}
