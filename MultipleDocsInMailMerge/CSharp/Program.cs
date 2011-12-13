//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExFor:MailMerge.Execute(DataRow)
//ExId:MultipleDocsInMailMerge
//ExSummary:Produce multiple documents during mail merge.
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Reflection;

using Aspose.Words;

namespace MultipleDocsInMailMerge
{
    class Program
    {
        public static void Main(string[] args)
        {
            //Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            ProduceMultipleDocuments(dataDir, "TestFile.doc");
        }

        public static void ProduceMultipleDocuments(string dataDir, string srcDoc)
        {
            // Open the database connection.
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataDir + "Customers.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();
            try
            {
                // Get data from a database.
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable data = new DataTable();
                da.Fill(data);

                // Open the template document.
                Document doc = new Document(dataDir + srcDoc);

                int counter = 1;
                // Loop though all records in the data source.
                foreach (DataRow row in data.Rows)
                {
                    // Clone the template instead of loading it from disk (for speed).
                    Document dstDoc = (Document)doc.Clone(true);

                    // Execute mail merge.
                    dstDoc.MailMerge.Execute(row);

                    // Save the document.
                    dstDoc.Save(string.Format(dataDir + "TestFile Out {0}.doc", counter++));
                }
            }
            finally
            {
                // Close the database.
                conn.Close();
            }
        }
    }
}
//ExEnd