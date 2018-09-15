using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeAndConditionalField
    {
        public static void Run()
        {
            // ExStart:MailMergeAndConditionalField           
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            // Open an existing document.
            Document doc = new Document(dataDir + "UnconditionalMergeFieldsAndRegions.docx");

            //Merge fields and merge regions are merged regardless of the parent IF field's condition.
            doc.MailMerge.UnconditionalMergeFieldsAndRegions = true;

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                new string[] { "FullName" },
                new object[] { "James Bond" });

            dataDir = dataDir + "UnconditionalMergeFieldsAndRegions_out.docx";
            doc.Save(dataDir);
            // ExEnd:MailMergeAndConditionalField
            Console.WriteLine("\nMail merge with conditional field has performed successfully.\nFile saved at " + dataDir);
        }
    }
}
