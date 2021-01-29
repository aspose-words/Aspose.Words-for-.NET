using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace _02._01_FindandReplaceTextinDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");

            // Replaces all 'sad' and 'mad' occurrences with 'bad'.
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = true
            };

            doc.Range.Replace("document", "document replaced", options);

            // Replaces all 'sad' and 'mad' occurrences with 'bad'.
            doc.Range.Replace(new Regex("[s|m]ad"), "bad");

            doc.Save("FindandReplaceTextinDocuments.docx");
        }
    }
}
