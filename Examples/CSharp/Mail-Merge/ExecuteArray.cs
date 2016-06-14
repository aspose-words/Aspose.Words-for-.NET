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
    class ExecuteArray
    {
        public static void Run()
        {
            //ExStart:ExecuteArray           
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
            // Open an existing document.
            Document doc = new Document(dataDir + "MailMerge.ExecuteArray.doc");

            // Trim trailing and leading whitespaces mail merge values
            doc.MailMerge.TrimWhitespaces = false;

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            dataDir = dataDir + "MailMerge.ExecuteArray_out_.doc";
            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
            doc.Save(dataDir);
            //ExEnd:ExecuteArray
            Console.WriteLine("\nSimple Mail merge performed with array data successfully.\nFile saved at " + dataDir);
        }
    }
}
