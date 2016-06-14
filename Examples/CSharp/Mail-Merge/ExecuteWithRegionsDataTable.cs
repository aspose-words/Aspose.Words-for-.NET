using Aspose.Words;
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
    class ExecuteWithRegionsDataTable
    {
        public static void Run()
        {
            //ExStart:ExecuteWithRegionsDataTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting();
            string fileName = "MailMerge.ExecuteWithRegions.doc";
            Document doc = new Document(dataDir + fileName);

            int orderId = 10444;

            // Perform several mail merge operations populating only part of the document each time.

            // Use DataTable as a data source.
            DataTable orderTable = GetTestOrder(orderId);
            doc.MailMerge.ExecuteWithRegions(orderTable);

            // Instead of using DataTable you can create a DataView for custom sort or filter and then mail merge.
            DataView orderDetailsView = new DataView(GetTestOrderDetails(orderId));
            orderDetailsView.Sort = "ExtendedPrice DESC";
            doc.MailMerge.ExecuteWithRegions(orderDetailsView);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            doc.Save(dataDir);
            //ExEnd:ExecuteWithRegionsDataTable

            Console.WriteLine("\nMail merge executed successfully with repeatable regions.\nFile saved at " + dataDir);
        }
        //ExStart:ExecuteWithRegionsDataTableMethods
        private static DataTable GetTestOrder(int orderId)
        {
            DataTable table = ExecuteDataTable(string.Format(
                "SELECT * FROM AsposeWordOrders WHERE OrderId = {0}", orderId));
            table.TableName = "Orders";
            return table;
        }
        private static DataTable GetTestOrderDetails(int orderId)
        {
            DataTable table = ExecuteDataTable(string.Format(
                "SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {0} ORDER BY ProductID", orderId));
            table.TableName = "OrderDetails";
            return table;
        }
        /// <summary>
        /// Utility function that creates a connection, command, 
        /// executes the command and return the result in a DataTable.
        /// </summary>
        private static DataTable ExecuteDataTable(string commandText)
        {
            // Open the database connection.
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                RunExamples.GetDataDir_Database() + "Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            // Create and execute a command.
            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable table = new DataTable();
            da.Fill(table);

            // Close the database.
            conn.Close();

            return table;
        }
        //ExEnd:ExecuteWithRegionsDataTableMethods
    }
}
