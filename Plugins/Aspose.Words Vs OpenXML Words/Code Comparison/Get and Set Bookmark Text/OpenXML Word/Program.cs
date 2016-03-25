using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName="Test.docx";
            IDictionary<String, BookmarkStart> bookmarkMap =
     new Dictionary<String, BookmarkStart>();
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(fileName, true))
            {
                foreach (BookmarkStart bookmarkStart in wordDocument.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                {
                    // foreach (BookmarkStart bookmarkStart in file.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
                    //{
                    bookmarkMap[bookmarkStart.Name] = bookmarkStart;

                    foreach (BookmarkStart bookmark in bookmarkMap.Values)
                    {
                        Run bookmarkText = bookmark.NextSibling<Run>();
                        if (bookmarkText != null)
                        {
                            bookmarkText.GetFirstChild<Text>().Text = "blah";
                        }
                    }

                }
            }
        }
    }
}
