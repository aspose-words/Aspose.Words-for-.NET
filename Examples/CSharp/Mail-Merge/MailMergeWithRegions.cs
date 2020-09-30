using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeWithRegionsExample
    {
        //ExStart:MailMergeWithRegions
        public void MailMergeWithRegions()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The start point of mail merge with regions the dataset.
            builder.InsertField(" MERGEFIELD TableStart:Customers");
            // Data from rows of the "CustomerName" column of the "Customers" table will go in this MERGEFIELD.
            builder.Write("Orders for ");
            builder.InsertField(" MERGEFIELD CustomerName");
            builder.Write(":");

            // Create column headers
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // We have a second data table called "Orders", which has a many-to-one relationship with "Customers"
            // picking up rows with the same CustomerID value.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.InsertField(" MERGEFIELD ItemName");
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Quantity");
            builder.InsertField(" MERGEFIELD TableEnd:Orders");
            builder.EndTable();

            // The end point of mail merge with regions.
            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            // Pass our dataset to perform mail merge with regions.          
            DataSet customersAndOrders = CreateDataSet();
            doc.MailMerge.ExecuteWithRegions(customersAndOrders);

            // Save the result
            doc.Save("Your local path to save the document" + "MailMerge.ExecuteWithRegions.docx");
        }
        //ExEnd:MailMergeWithRegions
        
        //ExStart:CreateDataSet
        private static DataSet CreateDataSet()
        {
            // Create the customers table.
            DataTable tableCustomers = new DataTable("Customers");
            tableCustomers.Columns.Add("CustomerID");
            tableCustomers.Columns.Add("CustomerName");
            tableCustomers.Rows.Add(new object[] { 1, "John Doe" });
            tableCustomers.Rows.Add(new object[] { 2, "Jane Doe" });

            // Create the orders table.
            DataTable tableOrders = new DataTable("Orders");
            tableOrders.Columns.Add("CustomerID");
            tableOrders.Columns.Add("ItemName");
            tableOrders.Columns.Add("Quantity");
            tableOrders.Rows.Add(new object[] { 1, "Hawaiian", 2 });
            tableOrders.Rows.Add(new object[] { 2, "Pepperoni", 1 });
            tableOrders.Rows.Add(new object[] { 2, "Chicago", 1 });

            // Add both tables to a data set.
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tableCustomers);
            dataSet.Tables.Add(tableOrders);

            // The "CustomerID" column, also the primary key of the customers table is the foreign key for the Orders table.
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"],    tableOrders.Columns["CustomerID"]);

            return dataSet;
        }
        //ExEnd:CreateDataSet

        public static void GetRegionsByName()
        {
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 

            //ExStart:GetRegionsByName
            Document doc = new Document(dataDir + "Mail merge regions.docx");

            IList<MailMergeRegionInfo> regions = doc.MailMerge.GetRegionsByName("Region1");
            Assert.AreEqual(1, doc.MailMerge.GetRegionsByName("Region1").Count);
            foreach (MailMergeRegionInfo region in regions) Assert.AreEqual("Region1", region.Name);

            regions = doc.MailMerge.GetRegionsByName("Region2");
            Assert.AreEqual(1, doc.MailMerge.GetRegionsByName("Region2").Count);
            foreach (MailMergeRegionInfo region in regions) Assert.AreEqual("Region2", region.Name);

            regions = doc.MailMerge.GetRegionsByName("NestedRegion1");
            Assert.AreEqual(2, doc.MailMerge.GetRegionsByName("NestedRegion1").Count);
            foreach (MailMergeRegionInfo region in regions) Assert.AreEqual("NestedRegion1", region.Name);
            //ExEnd:GetRegionsByName
        }
    }
}
