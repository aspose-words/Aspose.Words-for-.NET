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

            RemoveEmptyParagraphs(dataDir);
            RemoveUnusedFields(dataDir);
            RemoveContainingFields(dataDir);
            RemoveEmptyTableRows(dataDir);
            CleanupParagraphsWithPunctuationMarks(dataDir);
        }

        public static void RemoveEmptyParagraphs(string dataDir)
        {
            //ExStart:RemoveEmptyParagraphs
            Document doc = new Document(dataDir + "RemoveRowfromTable.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City" }, 
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
            
            doc.Save(dataDir + "MailMerge.ExecuteArray_out.doc");
            //ExEnd:RemoveEmptyParagraphs
        }

        public static void RemoveUnusedFields(string dataDir)
        {
            //ExStart:RemoveUnusedFields
            Document doc = new Document(dataDir + "RemoveRowfromTable.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City" }, 
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
            
            doc.Save(dataDir + "MailMerge.ExecuteArray_out.doc");
            //ExEnd:RemoveUnusedFields
        }

        public static void RemoveContainingFields(string dataDir)
        {
            //ExStart:RemoveContainingFields
            Document doc = new Document(dataDir + "RemoveRowfromTable.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields;

            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City" }, 
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
            
            doc.Save(dataDir + "MailMerge.ExecuteArray_out.doc");
            //ExEnd:RemoveContainingFields
        }

        public static void RemoveEmptyTableRows(string dataDir)
        {
            //ExStart:RemoveEmptyTableRows
            Document doc = new Document(dataDir + "RemoveRowfromTable.docx");

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyTableRows;

            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City" }, 
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
            
            doc.Save(dataDir + "MailMerge.ExecuteArray_out.doc");
            //ExEnd:RemoveEmptyTableRows
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
