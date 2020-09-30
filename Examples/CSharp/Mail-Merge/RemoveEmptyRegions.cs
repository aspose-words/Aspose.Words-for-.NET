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
    class RemoveEmptyRegions
    {
        public static void Run()
        {
            // ExStart:RemoveUnmergedRegions
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
            const string fileName = "TestFile Empty.doc";

            Document doc = new Document(dataDir + fileName);

            // Create an empty data source in the form of a DataSet containing no DataTable objects.
            DataSet data = new DataSet();

            // Enable the MailMergeCleanupOptions.RemoveUnusedRegions option.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;

            // Merge the data with the document by executing mail merge which will have no effect as there is no data.
            // However the regions found in the document will be removed automatically as they are unused.
            doc.MailMerge.ExecuteWithRegions(data);
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

            // Save the output document to disk.
            doc.Save(dataDir);
            // ExEnd:RemoveUnmergedRegions
            Debug.Assert(doc.MailMerge.GetFieldNames().Length == 0, "Error: There are still unused regions remaining in the document");

            Console.WriteLine("\nMail merge performed with empty regions successfully.\nFile saved at " + dataDir);
        }
    }
}
