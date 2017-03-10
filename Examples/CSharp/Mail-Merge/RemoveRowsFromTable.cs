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
    class RemoveRowsFromTable
    {
        public static void Run()
        {
            //Exstart:RemoveRowsFromTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            string fileName = "RemoveTableRows.doc";
            Document doc = new Document(dataDir + fileName);
            DataSet data = new DataSet();
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions | MailMergeCleanupOptions.RemoveEmptyTableRows;
            doc.MailMerge.MergeDuplicateRegions = true;
            doc.MailMerge.ExecuteWithRegions(data);
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the output document to disk.
            doc.Save(dataDir);
            //Exend:RemoveRowsFromTable
        }
    }
}
