using Aspose.Words.MailMerging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeCleanUp
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();

            CleanupParagraphsWithPunctuationMarks(dataDir);  
        }

        public static void CleanupParagraphsWithPunctuationMarks(string dataDir)
        {
            // ExStart:CleanupParagraphsWithPunctuationMarks
            string fileName = "MailMerge.CleanupPunctuationMarks.docx";
            // Open the document.
            Document doc = new Document(dataDir + fileName);

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            doc.MailMerge.CleanupParagraphsWithPunctuationMarks = false;

            doc.MailMerge.Execute(new string[] { "field1", "field2" }, new object[] { "", "" });

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the output document to disk.
            doc.Save(dataDir);
            // ExEnd:CleanupParagraphsWithPunctuationMarks

            Console.WriteLine("\nMail merge performed with cleanup paragraphs having punctuation marks successfully.\nFile saved at " + dataDir);
        }
    }
}
