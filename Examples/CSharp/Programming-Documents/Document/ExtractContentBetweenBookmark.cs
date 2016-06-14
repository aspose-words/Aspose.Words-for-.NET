using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractContentBetweenBookmark
    {
        public static void Run()
        {
            //ExStart:ExtractContentBetweenBookmark
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            Section section = doc.Sections[0];
            section.PageSetup.LeftMargin = 70.85;

            // Retrieve the bookmark from the document.
            Bookmark bookmark = doc.Range.Bookmarks["Bookmark1"];

            // We use the BookmarkStart and BookmarkEnd nodes as markers.
            BookmarkStart bookmarkStart = bookmark.BookmarkStart;
            BookmarkEnd bookmarkEnd = bookmark.BookmarkEnd;

            // Firstly extract the content between these nodes including the bookmark. 
            ArrayList extractedNodesInclusive = Common.ExtractContent(bookmarkStart, bookmarkEnd, true);
            Document dstDoc = Common.GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(dataDir + "TestFile.BookmarkInclusive_out_.doc");

            // Secondly extract the content between these nodes this time without including the bookmark.
            ArrayList extractedNodesExclusive = Common.ExtractContent(bookmarkStart, bookmarkEnd, false);
            dstDoc = Common.GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(dataDir + "TestFile.BookmarkExclusive_out_.doc");
            //ExEnd:ExtractContentBetweenBookmark
            Console.WriteLine("\nExtracted content between bookmarks successfully.\nFile saved at " + dataDir + "TestFile.BookmarkExclusive_out_.doc");
        }
    }
}
