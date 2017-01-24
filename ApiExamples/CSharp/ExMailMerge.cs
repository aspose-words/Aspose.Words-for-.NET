// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

//ExStart
//ExId:UsingReportingNamespace
//ExSummary:Include the following statement in your code if you are using mail merge functionality.
//ExEnd

using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Web;
using System.Collections;

using Aspose.Words.Fields;
using Aspose.Words;
using Aspose.Words.MailMerging;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExMailMerge : ApiExampleBase
    {
        [Test]
        public void ExecuteArray()
        {
            HttpResponse Response = null;

            //ExStart
            //ExFor:MailMerge.Execute(String[],Object[])
            //ExFor:ContentDisposition
            //ExFor:Document.Save(HttpResponse,String,ContentDisposition,SaveOptions)
            //ExId:MailMergeArray
            //ExSummary:Performs a simple insertion of data into merge fields and sends the document to the browser inline.
            // Open an existing document.
            Document doc = new Document(MyDir + "MailMerge.ExecuteArray.doc");

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                new string[] {"FullName", "Company", "Address", "Address2", "City"},
                new object[] {"James Bond", "MI5 Headquarters", "Milbank", "", "London"});

            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
            Assert.That(() => doc.Save(Response, @"\Artifacts\MailMerge.ExecuteArray.doc", ContentDisposition.Inline, null), Throws.TypeOf<ArgumentNullException>()); //Thrown because HttpResponse is null in the test.
            //ExEnd
        }

        [Test]
        public void ExecuteDataTable()
        {
            //ExStart
            //ExFor:Document
            //ExFor:MailMerge
            //ExFor:MailMerge.Execute(DataTable)
            //ExFor:Document.MailMerge
            //ExSummary:Executes mail merge from an ADO.NET DataTable.
            Document doc = new Document(MyDir + "MailMerge.ExecuteDataTable.doc");

            // This example creates a table, but you would normally load table from a database. 
            DataTable table = new DataTable("Test");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add(new object[] {"Thomas Hardy", "120 Hanover Sq., London"});
            table.Rows.Add(new object[] {"Paolo Accorti", "Via Monte Bianco 34, Torino"});

            // Field values from the table are inserted into the mail merge fields found in the document.
            doc.MailMerge.Execute(table);

            doc.Save(MyDir + @"\Artifacts\MailMerge.ExecuteDataTable.doc");
            //ExEnd
        }

        [Test]
        [TestCase(true, "first line\rsecond line\rthird line\f")]
        [TestCase(false, " first line\rsecond line\rthird line \f")]
        public void TrimWhiteSpaces(bool option, string expectedText)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertField("MERGEFIELD field", null);

            doc.MailMerge.TrimWhitespaces = option;
            doc.MailMerge.Execute(new[] { "field" }, new object[] { " first line\rsecond line\rthird line " });

            Assert.AreEqual(expectedText, doc.GetText());
        }

        [Test]
        public void ExecuteDataReader()
        {
            //ExStart
            //ExFor:MailMerge.Execute(IDataReader)
            //ExSummary:Executes mail merge from an ADO.NET DataReader.
            // Open the template document
            Document doc = new Document(MyDir + "MailingLabelsDemo.doc");

            // Open the database connection.
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + 
                DatabaseDir + "Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            

            // Open the data reader.
            OleDbCommand cmd = new OleDbCommand(
                "SELECT TOP 50 * FROM Customers ORDER BY Country, CompanyName", conn);
            OleDbDataReader dataReader = cmd.ExecuteReader();

            // Perform the mail merge
            doc.MailMerge.Execute(dataReader);

            // Close database.
            dataReader.Close();
            conn.Close();

            doc.Save(MyDir + @"\Artifacts\MailMerge.ExecuteDataReader.doc");
            //ExEnd
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void ExecuteDataViewCaller()
        {
            this.ExecuteDataView();
        }
        
        //ExStart
        //ExFor:MailMerge.Execute(DataView)
        //ExSummary:Executes mail merge from an ADO.NET DataView.
        public void ExecuteDataView()
        {
            // Open the document that we want to fill with data.
            Document doc = new Document(MyDir + "MailMerge.ExecuteDataView.doc");

            // Get the data from the database.
            DataTable orderTable = GetOrders();
            
            // Create a customized view of the data.
            DataView orderView = new DataView(orderTable);
            orderView.RowFilter = "OrderId = 10444";
            
            // Populate the document with the data.
            doc.MailMerge.Execute(orderView);

            doc.Save(MyDir + @"\Artifacts\MailMerge.ExecuteDataView.doc");
        }

        private static DataTable GetOrders()
        {
            // Open a database connection.
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + 
                DatabaseDir + "Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            // Create the command.
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM AsposeWordOrders", conn);

            // Fill an ADO.NET table from the command.
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable table = new DataTable();
            da.Fill(table);

            // Close database.
            conn.Close();

            return table;
        }
        //ExEnd


        [Test]
        public void ExecuteWithRegionsDataSet()
        {
            //ExStart
            //ExFor:MailMerge.ExecuteWithRegions(DataSet)
            //ExSummary:Executes a mail merge with repeatable regions from an ADO.NET DataSet.
            // Open the document. 
            // For a mail merge with repeatable regions, the document should have mail merge regions 
            // in the document designated with MERGEFIELD TableStart:MyTableName and TableEnd:MyTableName.
            Document doc = new Document(MyDir + "MailMerge.ExecuteWithRegions.doc");

            int orderId = 10444;

            // Populate tables and add them to the dataset.
            // For a mail merge with repeatable regions, DataTable.TableName should be 
            // set to match the name of the region defined in the document.
            DataSet dataSet = new DataSet();

            DataTable orderTable = GetTestOrder(orderId);
            dataSet.Tables.Add(orderTable);

            DataTable orderDetailsTable = GetTestOrderDetails(orderId);
            dataSet.Tables.Add(orderDetailsTable);

            // This looks through all mail merge regions inside the document and for each
            // region tries to find a DataTable with a matching name inside the DataSet.
            // If a table is found, its content is merged into the mail merge region in the document.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            doc.Save(MyDir + @"\Artifacts\MailMerge.ExecuteWithRegionsDataSet.doc");
            //ExEnd
        }


        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void ExecuteWithRegionsDataTableCaller()
        {
            this.ExecuteWithRegionsDataTable();
        }
        
        //ExStart
        //ExFor:Document.MailMerge
        //ExFor:MailMerge.ExecuteWithRegions(DataTable)
        //ExFor:MailMerge.ExecuteWithRegions(DataView)
        //ExId:MailMergeRegions
        //ExSummary:Executes a mail merge with repeatable regions.
        public void ExecuteWithRegionsDataTable()
        {
            Document doc = new Document(MyDir + "MailMerge.ExecuteWithRegions.doc");

            int orderId = 10444;

            // Perform several mail merge operations populating only part of the document each time.

            // Use DataTable as a data source.
            DataTable orderTable = GetTestOrder(orderId);
            doc.MailMerge.ExecuteWithRegions(orderTable);

            // Instead of using DataTable you can create a DataView for custom sort or filter and then mail merge.
            DataView orderDetailsView = new DataView(GetTestOrderDetails(orderId));
            orderDetailsView.Sort = "ExtendedPrice DESC";
            doc.MailMerge.ExecuteWithRegions(orderDetailsView);

            doc.Save(MyDir + @"\Artifacts\MailMerge.ExecuteWithRegionsDataTable.doc");
        }

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
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + 
                DatabaseDir + "Northwind.mdb";
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
        //ExEnd

        [Test]
        public void MappedDataFields()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:MailMerge.MappedDataFields
            //ExFor:MappedDataFieldCollection
            //ExFor:MappedDataFieldCollection.Add
            //ExId:MailMergeMappedDataFields
            //ExSummary:Shows how to add a mapping when a merge field in a document and a data field in a data source have different names.
            doc.MailMerge.MappedDataFields.Add("MyFieldName_InDocument", "MyFieldName_InDataSource");
            //ExEnd
        }

        [Test]
        public void GetFieldNames()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:MailMerge.GetFieldNames
            //ExId:MailMergeGetFieldNames
            //ExSummary:Shows how to get names of all merge fields in a document.
            string[] fieldNames = doc.MailMerge.GetFieldNames();
            //ExEnd
        }

        [Test]
        public void DeleteFields()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:MailMerge.DeleteFields
            //ExId:MailMergeDeleteFields
            //ExSummary:Shows how to delete all merge fields from a document without executing mail merge.
            doc.MailMerge.DeleteFields();
            //ExEnd
        }

        [Test]
        public void RemoveContainingFields()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:MailMerge.CleanupOptions
            //ExFor:MailMergeCleanupOptions
            //ExId:MailMergeRemoveContainingFields
            //ExSummary:Shows how to instruct the mail merge engine to remove any containing fields from around a merge field during mail merge.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields;
            //ExEnd
        }

        [Test]
        public void RemoveUnusedFields()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:MailMerge.CleanupOptions
            //ExFor:MailMergeCleanupOptions
            //ExId:MailMergeRemoveUnusedFields
            //ExSummary:Shows how to automatically remove unmerged merge fields during mail merge.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
            //ExEnd
        }

        [Test]
        public void RemoveEmptyParagraphs()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:MailMerge.CleanupOptions
            //ExFor:MailMergeCleanupOptions
            //ExId:MailMergeRemoveEmptyParagraphs
            //ExSummary:Shows how to make sure empty paragraphs that result from merging fields with no data are removed from the document.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            //ExEnd
        }

        [Test]
        public void UseNonMergeFields()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:MailMerge.UseNonMergeFields
            //ExSummary:Shows how to perform mail merge into merge fields and into additional fields types.
            doc.MailMerge.UseNonMergeFields = true;
            //ExEnd
        }

        [Test]
        [TestCase(true, "{{ testfield1 }}value 1{{ testfield3 }}\f")]
        [TestCase(false, "\u0013MERGEFIELD \"testfield1\"\u0014«testfield1»\u0015value 1\u0013MERGEFIELD \"testfield3\"\u0014«testfield3»\u0015\f")]
        public void MustasheTemplateSyntax(bool restoreTags, string sectionText)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("{{ testfield1 }}");
            builder.Write("{{ testfield2 }}");
            builder.Write("{{ testfield3 }}");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.PreserveUnusedTags = restoreTags;

            DataTable table = new DataTable("Test");
            table.Columns.Add("testfield2");
            table.Rows.Add(new object[] { "value 1" });

            doc.MailMerge.Execute(table);

            string paraText = DocumentHelper.GetParagraphText(doc, 0);

            Assert.AreEqual(sectionText, paraText);
        }

        [Test]
        public void TestMailMergeGetRegionsHierarchy()
        {
            //ExStart
            //ExFor:MailMerge.GetRegionsHierarchy
            //ExFor:MailMergeRegionInfo.Regions
            //ExFor:MailMergeRegionInfo.Name
            //ExFor:MailMergeRegionInfo.Fields
            //ExFor:MailMergeRegionInfo.StartField
            //ExFor:MailMergeRegionInfo.EndField
            //ExFor:MailMergeRegionInfo.Level
            //ExSummary:Shows how to get MailMergeRegionInfo and work with it
            Document doc = new Document(MyDir+ "MailMerge.TestRegionsHierarchy.doc");

            //Returns a full hierarchy of regions (with fields) available in the document.
            MailMergeRegionInfo regionInfo = doc.MailMerge.GetRegionsHierarchy();

            //Get top regions in the document
            ArrayList topRegions = regionInfo.Regions;
            Assert.AreEqual(2, topRegions.Count);
            Assert.AreEqual(((MailMergeRegionInfo)topRegions[0]).Name, "Region1");
            Assert.AreEqual(((MailMergeRegionInfo)topRegions[1]).Name, "Region2");
            Assert.AreEqual(1, ((MailMergeRegionInfo)topRegions[0]).Level);
            Assert.AreEqual(1, ((MailMergeRegionInfo)topRegions[1]).Level);

            //Get nested region in first top region
            ArrayList nestedRegions = ((MailMergeRegionInfo)topRegions[0]).Regions;
            Assert.AreEqual(2, nestedRegions.Count);
            Assert.AreEqual(((MailMergeRegionInfo)nestedRegions[0]).Name, "NestedRegion1");
            Assert.AreEqual(((MailMergeRegionInfo)nestedRegions[1]).Name, "NestedRegion2");
            Assert.AreEqual(2, ((MailMergeRegionInfo)nestedRegions[0]).Level);
            Assert.AreEqual(2, ((MailMergeRegionInfo)nestedRegions[1]).Level);

            //Get field list in first top region
            ArrayList fieldList = ((MailMergeRegionInfo)topRegions[0]).Fields;
            Assert.AreEqual(4, fieldList.Count);

            FieldMergeField startFieldMergeField = ((MailMergeRegionInfo)nestedRegions[0]).StartField;
            Assert.AreEqual("TableStart:NestedRegion1", startFieldMergeField.FieldName);

            FieldMergeField endFieldMergeField = ((MailMergeRegionInfo)nestedRegions[0]).EndField;
            Assert.AreEqual("TableEnd:NestedRegion1", endFieldMergeField.FieldName);
            //ExEnd
        }

        [Test]
        public void TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption()
        {
            Document document = new Document();
            document.MailMerge.UseNonMergeFields = true;
           
            MailMergeCallbackStub mailMergeCallbackStub = new MailMergeCallbackStub();
            document.MailMerge.MailMergeCallback = mailMergeCallbackStub;

            document.MailMerge.Execute(
                new string[0],
                new object[0]);

            Assert.AreEqual(1, mailMergeCallbackStub.TagsReplacedCounter);
        }

        [Test]
        [TestCase("Region1")]
        [TestCase("NestedRegion1")]
        public void GetRegionsByName(string regionName)
        {
            Document doc = new Document(MyDir + "MailMerge.RegionsByName.doc");

            ArrayList regions = doc.MailMerge.GetRegionsByName(regionName);
            Assert.AreEqual(2, regions.Count);

            foreach (MailMergeRegionInfo region in regions)
            {
                Assert.AreEqual(regionName, region.Name);
            }
        }

        private class MailMergeCallbackStub : IMailMergeCallback
        {
            public void TagsReplaced()
            {
                mTagsReplacedCounter++;
            }

            public int TagsReplacedCounter
            {
                get { return mTagsReplacedCounter; }
            }

            private int mTagsReplacedCounter;
        }
    }
}
