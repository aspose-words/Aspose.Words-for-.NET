using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MultipleDocsInMailMerge
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
            
            // Open the database connection.
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataDir + "Customers.mdb";
            OleDbConnection conn = new OleDbConnection(connString);

            try
            {
                conn.Open();

                // Get data from a database.
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable data = new DataTable();
                da.Fill(data);

                // Open the template document.
                Document doc = new Document(dataDir + "TestFile.Multiple Pages.doc");

                int counter = 1;
                // Loop though all records in the data source.
                foreach (DataRow row in data.Rows)
                {
                    // Clone the template instead of loading it from disk (for speed).
                    Document dstDoc = (Document)doc.Clone(true);

                    // Execute mail merge.
                    dstDoc.MailMerge.Execute(row);

                    // Save the document.
                    dstDoc.Save(string.Format(dataDir + "TestFile.Multiple Pages_out_ {0}.doc", counter++));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Close the database.
                conn.Close();
            }

            Console.WriteLine("\nMail merge performed and created multiple pages successfully.\nFile saved at " + dataDir + "TestFile.Multiple Pages_out_.doc");
        }
    }
}
