using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace _02._01_FindandReplaceTextinDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document("../../data/document.doc");

            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false, FindWholeWordsOnly = true
            };

            // Replace all occurrences of 'sad' and 'mad' with 'bad'.
            doc.Range.Replace("document", "document replaced", options);

            // Replace all occurrences of 'sad' and 'mad' found using a regex pattern with 'bad'.
            doc.Range.Replace(new Regex("[s|m]ad"), "bad");

            doc.Save("FindandReplaceTextinDocuments.docx");
        }
    }
}
