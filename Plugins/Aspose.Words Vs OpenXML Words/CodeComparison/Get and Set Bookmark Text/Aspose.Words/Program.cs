using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "Test.docx";
            Document doc = new Document(fileName);

            // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

            // Get the name and text of the bookmark.
            string name = bookmark.Name;
            string text = bookmark.Text;

            // Set the name and text of the bookmark.
            bookmark.Name = "RenamedBookmark";
            bookmark.Text = "This is a new bookmarked text.";
            doc.Save(fileName);
        }
    }
}
