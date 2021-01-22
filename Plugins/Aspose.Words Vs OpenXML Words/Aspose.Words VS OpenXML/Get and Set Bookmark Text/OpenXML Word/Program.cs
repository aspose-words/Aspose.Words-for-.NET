// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string File = FilePath + "Get and Set Bookmark Text - OpenXML.docx";

            IDictionary<String, BookmarkStart> bookmarkMap = new Dictionary<String, BookmarkStart>();

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(File, true))
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
