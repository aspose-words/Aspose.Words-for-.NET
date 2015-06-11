//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenBookmark
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_WorkingWithDocument();

            Document doc = new Document(dataDir + "TestFile.doc");

            // Retrieve the bookmark from the document.
            Aspose.Words.Bookmark bookmark = doc.Range.Bookmarks["Bookmark1"];

            // We use the BookmarkStart and BookmarkEnd nodes as markers.
            BookmarkStart bookmarkStart = bookmark.BookmarkStart;
            BookmarkEnd bookmarkEnd = bookmark.BookmarkEnd;

            // Firstly extract the content between these nodes including the bookmark. 
            ArrayList extractedNodesInclusive = Common.ExtractContent(bookmarkStart, bookmarkEnd, true);
            Document dstDoc = Common.GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(dataDir + "TestFile.BookmarkInclusive Out.doc");

            // Secondly extract the content between these nodes this time without including the bookmark.
            ArrayList extractedNodesExclusive = Common.ExtractContent(bookmarkStart, bookmarkEnd, false);
            dstDoc = Common.GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(dataDir + "TestFile.BookmarkExclusive Out.doc");

            Console.WriteLine("\nExtracted content between bookmarks successfully.\nFile saved at " + dataDir + "TestFile.BookmarkExclusive Out.doc");
        }
    }
}
