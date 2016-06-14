using System;
using System.IO;
using System.Reflection;
using Aspose.Words.Tables;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Bookmarks
{
    class AccessBookmarks
    {
        public static void Run()
        {
            //ExStart:AccessBookmarks
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithBookmarks();

            Document doc = new Document(dataDir + "Bookmarks.doc");

            // By index.
            Bookmark bookmark1 = doc.Range.Bookmarks[0];
           
            // By name.
            Bookmark bookmark2 = doc.Range.Bookmarks["Bookmark2"];
            //ExEnd:AccessBookmarks
            Console.WriteLine("\nBookmark by name is " + bookmark1.Name + " and bookmark by index is " + bookmark2.Name);
        }
        
    }
}
