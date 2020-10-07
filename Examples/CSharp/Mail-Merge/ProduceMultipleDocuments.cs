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
            //Put the path to the documents directory and open the template:
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            Document doc = new Document(dataDir + "TestFile.doc");

            // Open the database connection.
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataDir + "Customers.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            // Get data from a database.
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable data = new DataTable();
            da.Fill(data);

            //Perform a loop through each DataRow to iterate through the DataTable.
            //Clone the template document instead of loading it from disk for better speed performance before the mail merge operation.
            //You can load the template document from a file or stream but it is faster to load the document only once and then clone it in memory before each mail merge operation.

            int counter = 1;
            foreach (DataRow row in data.Rows)
            {
                Document dstDoc = (Document) doc.Clone(true);
                dstDoc.MailMerge.Execute(row);
                dstDoc.Save(string.Format(dataDir + "TestFile_out{0}.doc", counter++));
            }

            Console.WriteLine("\nProduce multiple documents performed successfully.\nFile saved at " + dataDir);
            //ExEnd:ProduceMultipleDocuments           
        }
    }
}
