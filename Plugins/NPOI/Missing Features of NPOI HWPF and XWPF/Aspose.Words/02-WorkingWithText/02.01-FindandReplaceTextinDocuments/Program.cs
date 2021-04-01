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
            // Check for an Aspose.Words license file in the local file system and apply it, if it exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                Aspose.Words.License license = new Aspose.Words.License();

                // Use the license from the bin/debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false, FindWholeWordsOnly = true
            };

            builder.Writeln("This document will be saved as a .doc file.");

            // Replace all occurences of ".doc" with ".docx".
            doc.Sections[0].Range.Replace(".doc", ".docx", options);

            builder.Writeln("Sad, mad.");

            // Replace all occurrences of 'sad' and 'mad' found using a regex pattern with 'bad'.
            doc.Sections[1].Range.Replace(new Regex("[s|m]ad"), "bad");

            doc.Save("FindandReplaceTextinDocuments.docx");
        }
    }
}
