using Aspose.Words;
using Aspose.Words.MailMerging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class SimpleMailMerge
    {
        // ExStart:SimpleMailMergeExecuteArray 
        public static void SimpleMailMergeExecuteArray()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 

            // Include the code for our template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create Merge Fields.
            builder.InsertField(" MERGEFIELD CustomerName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Item ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Quantity ");

            builder.Document.Save(dataDir + "MailMerge.TestTemplate.docx");
            
            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(new string[] { "CustomerName", "Item", "Quantity" },
                new object[] { "John Doe", "Hawaiian", "2" });

            builder.Document.Save(dataDir + "MailMerge.Simple.docx");
            // ExEnd:SimpleMailMergeExecuteArray
            Console.WriteLine("\nSimple Mail merge performed with array data successfully.\nFile saved at " + dataDir);
        }
    }
}
