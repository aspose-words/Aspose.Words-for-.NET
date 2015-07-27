using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;

namespace _02._01_FindandReplaceTextinDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");

            // Replaces all 'sad' and 'mad' occurrences with 'bad'
            doc.Range.Replace("document", "document replaced", false, true);

            // Replaces all 'sad' and 'mad' occurrences with 'bad'
            doc.Range.Replace(new Regex("[s|m]ad"), "bad");

            doc.Save("replacedDocument.doc");
        }
    }
}
