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
using System.Web;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class ProduceMultipleDocuments
    {
        public static void Run()
        {
            //ExStart:ProduceMultipleDocuments            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            // Open the database connection.
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataDir + "Customers.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();
            // Get data from a database.
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable data = new DataTable();
            da.Fill(data);

            // Open the template document.
            Document doc = new Document(dataDir + "TestFile.doc");

            int counter = 1;
            // Loop though all records in the data source.
            foreach (DataRow row in data.Rows)
            {
                // Clone the template instead of loading it from disk (for speed).
                Document dstDoc = (Document)doc.Clone(true);

                // Execute mail merge.
                dstDoc.MailMerge.Execute(row);

                // Save the document.
                dstDoc.Save(string.Format(dataDir + "TestFile_out_{0}.doc", counter++));
            }
            //ExEnd:ProduceMultipleDocuments
            Console.WriteLine("\nProduce multiple documents performed successfully.\nFile saved at " + dataDir);            
        }
    }
}
