using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Tables;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class BookmarkNameAndText
    {
        public static void Run()
        {
            //ExStart:BookmarkNameAndText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();

            Document doc = new Document(dataDir + "Bookmark.doc");

            // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

            // Get the name and text of the bookmark.
            string name = bookmark.Name;
            string text = bookmark.Text;

            // Set the name and text of the bookmark.
            bookmark.Name = "RenamedBookmark";
            bookmark.Text = "This is a new bookmarked text.";
            //ExEnd:BookmarkNameAndText
            Console.WriteLine("\nBookmark text and name get and set successfully." + dataDir);
        }
        
    }
}
