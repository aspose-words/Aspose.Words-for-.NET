using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using Aspose.Words;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace DocsExamples.Mail_Merge_and_Reporting
{
    internal class BaseOperations : DocsExamplesBase
    {
        [Test]
        public void SimpleMailMerge()
        {
            //ExStart:SimpleMailMerge
            // Include the code for our template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create Merge Fields.
            builder.InsertField(" MERGEFIELD CustomerName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Item ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Quantity ");

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(new string[] { "CustomerName", "Item", "Quantity" },
                new object[] { "John Doe", "Hawaiian", "2" });

            doc.Save(ArtifactsDir + "BaseOperations.SimpleMailMerge.docx");
            //ExEnd:SimpleMailMerge
        }

        [Test]
        public void UseIfElseMustache()
        {
            //ExStart:UseOfifelseMustacheSyntax
            Document doc = new Document(MyDir + "Mail merge destinations - Mustache syntax.docx");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.Execute(new[] { "GENDER" }, new object[] { "MALE" });

            doc.Save(ArtifactsDir + "BaseOperations.IfElseMustache.docx");
            //ExEnd:UseOfifelseMustacheSyntax
        }

        [Test]
        public void MustacheSyntaxUsingDataTable()
        {
            //ExStart:MustacheSyntaxUsingDataTable
            Document doc = new Document(MyDir + "Mail merge destinations - Vendor.docx");

            // Loop through each row and fill it with data.
            DataTable dataTable = new DataTable("list");
            dataTable.Columns.Add("Number");
            for (int i = 0; i < 10; i++)
            {
                DataRow dataRow = dataTable.NewRow();
                dataTable.Rows.Add(dataRow);
                dataRow[0] = "Number " + i;
            }

            // Activate performing a mail merge operation into additional field types.
            doc.MailMerge.UseNonMergeFields = true;

            doc.MailMerge.ExecuteWithRegions(dataTable);

            doc.Save(ArtifactsDir + "WorkingWithXmlData.MustacheSyntaxUsingDataTable.docx");
            //ExEnd:MustacheSyntaxUsingDataTable
        }

        [Test]
        public void ExecuteWithRegionsDataTable()
        {
            //ExStart:ExecuteWithRegionsDataTable
            Document doc = new Document(MyDir + "Mail merge destinations - Orders.docx");

            // Use DataTable as a data source.
            int orderId = 10444;
            DataTable orderTable = GetTestOrder(orderId);
            doc.MailMerge.ExecuteWithRegions(orderTable);

            // Instead of using DataTable, you can create a DataView for custom sort or filter and then mail merge.
            DataView orderDetailsView = new DataView(GetTestOrderDetails(orderId));
            orderDetailsView.Sort = "ExtendedPrice DESC";

            // Execute the mail merge operation.
            doc.MailMerge.ExecuteWithRegions(orderDetailsView);

            doc.Save(ArtifactsDir + "MailMerge.ExecuteWithRegions.docx");
            //ExEnd:ExecuteWithRegionsDataTable
        }

        //ExStart:ExecuteWithRegionsDataTableMethods
        private DataTable GetTestOrder(int orderId)
        {
            DataTable table = ExecuteDataTable($"SELECT * FROM AsposeWordOrders WHERE OrderId = {orderId}");
            table.TableName = "Orders";
            
            return table;
        }

        private DataTable GetTestOrderDetails(int orderId)
        {
            DataTable table = ExecuteDataTable(
                $"SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {orderId} ORDER BY ProductID");
            table.TableName = "OrderDetails";
            
            return table;
        }

        /// <summary>
        /// Utility function that creates a connection, command, executes the command and returns the result in a DataTable.
        /// </summary>
        private DataTable ExecuteDataTable(string commandText)
        {
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseDir + "Northwind.mdb";

            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand(commandText, conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);

            DataTable table = new DataTable();
            da.Fill(table);

            conn.Close();

            return table;
        }
        //ExEnd:ExecuteWithRegionsDataTableMethods

        [Test]
        public void ProduceMultipleDocuments()
        {
            //ExStart:ProduceMultipleDocuments
            string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseDir + "Northwind.mdb";

            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");

            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();
            
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            
            DataTable data = new DataTable();
            da.Fill(data);

            // Perform a loop through each DataRow to iterate through the DataTable. Clone the template document
            // instead of loading it from disk for better speed performance before the mail merge operation.
            // You can load the template document from a file or stream but it is faster to load the document
            // only once and then clone it in memory before each mail merge operation.
            
            int counter = 1;
            foreach (DataRow row in data.Rows)
            {
                Document dstDoc = (Document) doc.Clone(true);

                dstDoc.MailMerge.Execute(row);

                dstDoc.Save(string.Format(ArtifactsDir + "BaseOperations.ProduceMultipleDocuments_{0}.docx", counter++));
            }
            //ExEnd:ProduceMultipleDocuments
        }

        //ExStart:MailMergeWithRegions
        [Test]
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

            // Create column headers.
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

            doc.Save(ArtifactsDir + "BaseOperations.MailMergeWithRegions.docx");
        }
        //ExEnd:MailMergeWithRegions

        //ExStart:CreateDataSet
        private DataSet CreateDataSet()
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
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"], tableOrders.Columns["CustomerID"]);

            return dataSet;
        }
        //ExEnd:CreateDataSet

        [Test]
        public void GetRegionsByName()
        {
            //ExStart:GetRegionsByName
            Document doc = new Document(MyDir + "Mail merge regions.docx");

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