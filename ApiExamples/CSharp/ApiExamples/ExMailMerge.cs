// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.SqlClient;
using Aspose.Words.Fields;
using Aspose.Words;
using Aspose.Words.MailMerging;
using NUnit.Framework;
using NUnit.Framework.Constraints;
#if !(NETSTANDARD2_0 || __MOBILE__ || MAC)
using System.Data.OleDb;
using System.Web;

#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExMailMerge : ApiExampleBase
    {
#if !(NETSTANDARD2_0 || __MOBILE__ || MAC)
        [Test]
        public void ExecuteArray()
        {
            HttpResponse Response = null;

            //ExStart
            //ExFor:MailMerge.Execute(String[], Object[])
            //ExFor:ContentDisposition
            //ExFor:Document.Save(HttpResponse,String,ContentDisposition,SaveOptions)
            //ExId:MailMergeArray
            //ExSummary:Performs a simple insertion of data into merge fields and sends the document to the browser inline.
            // Open an existing document.
            Document doc = new Document(MyDir + "MailMerge.ExecuteArray.doc");

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(new String[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
            Assert.That(() => doc.Save(Response, "Artifacts/MailMerge.ExecuteArray.doc", ContentDisposition.Inline, null), Throws.TypeOf<ArgumentNullException>()); //Thrown because HttpResponse is null in the test.
            //ExEnd
        }
#endif

        [Test]
        public void ExecuteDataTable()
        {
            //ExStart
            //ExFor:Document
            //ExFor:MailMerge
            //ExFor:MailMerge.Execute(DataTable)
            //ExFor:MailMerge.Execute(DataRow)
            //ExFor:Document.MailMerge
            //ExSummary:Executes mail merge from an ADO.NET DataTable.
            Document doc = new Document(MyDir + "MailMerge.ExecuteDataTable.doc");

            // This example creates a table, but you would normally load table from a database. 
            DataTable table = new DataTable("Test");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London" });
            table.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

            // Field values from the table are inserted into the mail merge fields found in the document.
            doc.MailMerge.Execute(table);

            doc.Save(ArtifactsDir + "MailMerge.ExecuteDataTable.doc");

            // Open a fresh copy of our document to perform another mail merge.
            doc = new Document(MyDir + "MailMerge.ExecuteDataTable.doc");

            // We can also source values for a mail merge from a single row in the table
            doc.MailMerge.Execute(table.Rows[1]);

            doc.Save(ArtifactsDir + "MailMerge.ExecuteDataTable.OneRow.doc");
            //ExEnd
        }

        [Test]
        public void ExecuteDataView()
        {
            //ExStart
            //ExFor:MailMerge.Execute(DataView)
            //ExSummary:Shows how to process a DataTable's data with a DataView before using it in a mail merge.
            // Create a new document and populate it with merge fields
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Congratulations ");
            builder.InsertField(" MERGEFIELD Name");
            builder.Write(" for passing with a grade of ");
            builder.InsertField(" MERGEFIELD Grade");

            // Create a data table that merge data will be sourced from 
            DataTable table = new DataTable("ExamResults");
            table.Columns.Add("Name");
            table.Columns.Add("Grade");
            table.Rows.Add(new object[] { "John Doe", "67" });
            table.Rows.Add(new object[] { "Jane Doe", "81" });
            table.Rows.Add(new object[] { "John Cardholder", "47" });
            table.Rows.Add(new object[] { "Joe Bloggs", "75" });

            // If we execute the mail merge on the table, a page will be created for each row in the order that it appears in the table
            // If we want to sort/filter rows without changing the table, we can use a data view
            DataView view = new DataView(table);
            view.Sort = "Grade DESC";
            view.RowFilter = "Grade >= 50";

            // This mail merge will be executed on a view where the rows are sorted by the "Grade" column
            // and rows where the Grade values are below 50 are filtered out
            doc.MailMerge.Execute(view);

            doc.Save(ArtifactsDir + "MailMerge.ExecuteDataView.docx");
            //ExEnd
        }

        [Test]
        public void ExecuteDataReader()
        {
            //ExStart
            //ExFor:MailMerge.Execute(IDataReader)
            //ExSummary:Shows how to use a Data Reader to execute a mail merge on a database.
            // Create a new document and populate it with merge fields
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Product:\t");
            builder.InsertField(" MERGEFIELD ProductName");
            builder.Write("\nSupplier:\t");
            builder.InsertField(" MERGEFIELD CompanyName");
            builder.Writeln();
            builder.InsertField(" MERGEFIELD QuantityPerUnit");
            builder.Write(" for $");
            builder.InsertField(" MERGEFIELD UnitPrice");

            // Create a connection string which points to the "Northwind" database file in our local file system and open a connection
            string connectionString = @"Driver={Microsoft Access Driver (*.mdb)};Dbq=" + DatabaseDir + "Northwind.mdb";

            using (OdbcConnection connection = new OdbcConnection())
            {
                connection.ConnectionString = connectionString;
                connection.Open();

                // Create an SQL command that will source data for our mail merge
                // The command has to comply to the driver we are using, which in this case is "ODBC"
                // The names of the columns returned by this SELECT statement should correspond to the merge fields we placed above
                OdbcCommand command = connection.CreateCommand();
                command.CommandText = @"SELECT Products.ProductName, Suppliers.CompanyName, Products.QuantityPerUnit, {fn ROUND(Products.UnitPrice,2)} as UnitPrice
                                        FROM Products 
                                        INNER JOIN Suppliers 
                                        ON Products.SupplierID = Suppliers.SupplierID";

                // This will run the command and store the data in the reader
                OdbcDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

                // Now we can take the data from the reader and use it in the mail merge
                doc.MailMerge.Execute(reader);
            }

            doc.Save(ArtifactsDir + "MailMerge.ExecuteDataReader.docx");
            //ExEnd
        }

        [Test]
        public void ExecuteADO()
        {
            //ExStart
            //ExFor:MailMerge.ExecuteADO(Object)
            //ExSummary:Shows how to run a mail merge on an ADO dataset
            // Create a blank document and populate it with MERGEFIELDS that will accept data when a mail merge is executed
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Product:\t");
            builder.InsertField(" MERGEFIELD ProductName");
            builder.Writeln();
            builder.InsertField(" MERGEFIELD QuantityPerUnit");
            builder.Write(" for $");
            builder.InsertField(" MERGEFIELD UnitPrice");

            // To work with ADO DataSets, we need to add a reference to the Microsoft ActiveX Data Objects library,
            // which is included in the .NET distribution and stored in "adodb.dll", then create a connection
            ADODB.Connection connection = new ADODB.Connection();

            // We will then create a connection string which points to the "Northwind" database file in our local file system and open a connection
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseDir + "Northwind.mdb";
            connection.Open(connectionString);

            // Create a record set
            ADODB.Recordset recordset = new ADODB.Recordset();

            // Run an SQL command on the database we are connected to to populate our dataset
            // The names of the columns returned here correspond to the values of the MERGEFIELDS that will accomodate our data
            string command = @"SELECT ProductName, QuantityPerUnit, UnitPrice FROM Products";
            recordset.Open(command, connection);

            // Execute the mail merge and save the document
            doc.MailMerge.ExecuteADO(recordset);
            doc.Save(ArtifactsDir + "MailMerge.ExecuteADO.docx");
            //ExEnd
        }

        //ExStart
        //ExFor:MailMerge.ExecuteWithRegions(System.Data.DataSet)
        //ExSummary:Shows how to create a nested mail merge with regions with data from a data set with two related tables.
        [Test]
        public void ExecuteWithRegionsNested()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a MERGEFIELD with a value of "TableStart:Customers"
            // Normally, MERGEFIELDs specify the name of the column that they take row data from
            // "TableStart:Customers" however means that we are starting a mail merge region which belongs to a table called "Customers"
            // This will start the outer region and an "TableEnd:Customers" MERGEFIELD will signify its end 
            builder.InsertField(" MERGEFIELD TableStart:Customers");

            // Data from rows of the "CustomerName" column of the "Customers" table will go in this MERGEFIELD
            builder.Write("Orders for ");
            builder.InsertField(" MERGEFIELD CustomerName");
            builder.Write(":");

            // Create column headers for a table which will contain values from the second inner region
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // We have a second data table called "Orders", which has a many-to-one relationship with "Customers", related by a "CustomerID" column
            // We will start this inner mail merge region over which the "Orders" table will preside,
            // which will iterate over the "Orders" table once for each merge of the outer "Customers" region, picking up rows with the same CustomerID value
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.InsertField(" MERGEFIELD ItemName");
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Quantity");

            // End the inner region
            // One stipulation of using regions and tables is that the opening and closing of a mail merge region must only happen over one row of a document's table  
            builder.InsertField(" MERGEFIELD TableEnd:Orders");
            builder.EndTable();

            // End the outer region
            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            DataSet customersAndOrders = CreateDataSet();
            doc.MailMerge.ExecuteWithRegions(customersAndOrders);

            doc.Save(ArtifactsDir + "MailMerge.ExecuteWithRegionsNested.docx");
        }

        /// <summary>
        /// Generates a data set which has two data tables named "Customers" and "Orders",
        /// with a one-to-many relationship between the former and latter on the "CustomerID" column
        /// </summary>
        private DataSet CreateDataSet()
        {
            // Create the outer mail merge
            DataTable tableCustomers = new DataTable("Customers");
            tableCustomers.Columns.Add("CustomerID");
            tableCustomers.Columns.Add("CustomerName");
            tableCustomers.Rows.Add(new object[] { 1, "John Doe" });
            tableCustomers.Rows.Add(new object[] { 2, "Jane Doe" });

            // Create the table for the inner merge
            DataTable tableOrders = new DataTable("Orders");
            tableOrders.Columns.Add("CustomerID");
            tableOrders.Columns.Add("ItemName");
            tableOrders.Columns.Add("Quantity");
            tableOrders.Rows.Add(new object[] { 1, "Hawaiian", 2 });
            tableOrders.Rows.Add(new object[] { 2, "Pepperoni", 1 });
            tableOrders.Rows.Add(new object[] { 2, "Chicago", 1 });

            // Add both tables to a data set
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tableCustomers);
            dataSet.Tables.Add(tableOrders);

            // The "CustomerID" column, also the primary key of the customers table is the foreign key for the Orders table
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"], tableOrders.Columns["CustomerID"]);

            return dataSet;
        }
        //ExEnd

        [Test]
        public void ExecuteWithRegionsConcurrent()
        {
            //ExStart
            //ExFor:MailMerge.ExecuteWithRegions(DataTable)
            //ExFor:MailMerge.ExecuteWithRegions(DataView)
            //ExSummary:Shows how to use regions to execute two separate mail merges in one document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // If we want to perform two consecutive mail merges on one document while taking data from two tables
            // that are related to each other in any way, we can separate the mail merges with regions
            // A mail merge region starts and ends with "TableStart:[RegionName]" and "TableEnd:[RegionName]" MERGEFIELDs
            // These regions are separate for unrelated data, while they can be nested for hierarchical data
            builder.Writeln("\tCities: ");
            builder.InsertField(" MERGEFIELD TableStart:Cities");
            builder.InsertField(" MERGEFIELD Name");
            builder.InsertField(" MERGEFIELD TableEnd:Cities");
            builder.InsertParagraph();

            // Both MERGEFIELDs refer to a same column name, but values for each will come from different data tables
            builder.Writeln("\tFruit: ");
            builder.InsertField(" MERGEFIELD TableStart:Fruit");
            builder.InsertField(" MERGEFIELD Name");
            builder.InsertField(" MERGEFIELD TableEnd:Fruit");

            // Create two data tables that aren't linked or related in any way which we still want in the same document
            DataTable tableCities = new DataTable("Cities");
            tableCities.Columns.Add("Name");
            tableCities.Rows.Add(new object[] { "Washington" });
            tableCities.Rows.Add(new object[] { "London" });
            tableCities.Rows.Add(new object[] { "New York" });

            DataTable tableFruit = new DataTable("Fruit");
            tableFruit.Columns.Add("Name");
            tableFruit.Rows.Add(new object[] { "Cherry"});
            tableFruit.Rows.Add(new object[] { "Apple" });
            tableFruit.Rows.Add(new object[] { "Watermelon" });
            tableFruit.Rows.Add(new object[] { "Banana" });

            // We will need to run one mail merge per table
            // This mail merge will populate the MERGEFIELDs in the "Cities" range, while leaving the fields in "Fruit" empty
            doc.MailMerge.ExecuteWithRegions(tableCities);

            // Run a second merge for the "Fruit" table
            // We can use a DataView to sort or filter values of a DataTable before it is merged
            DataView dv = new DataView(tableFruit);
            dv.Sort = "Name ASC";
            doc.MailMerge.ExecuteWithRegions(dv);

            doc.Save(ArtifactsDir + "MailMerge.ExecuteWithRegionsConcurrent.docx");
            //ExEnd
        }

        [Test]
        public void TrimWhiteSpaces()
        {
            //ExStart
            //ExFor:MailMerge.TrimWhitespaces
            //ExSummary:Shows how to trimmed whitespaces from mail merge values.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("MERGEFIELD field", null);

            doc.MailMerge.TrimWhitespaces = true;
            doc.MailMerge.Execute(new[] { "field" }, new object[] { " first line\rsecond line\rthird line " });

            Assert.AreEqual("first line\rsecond line\rthird line\f", doc.GetText());
            //ExEnd
        }

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
        public void MailMergeGetFieldNames()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:MailMerge.GetFieldNames
            //ExId:MailMergeGetFieldNames
            //ExSummary:Shows how to get names of all merge fields in a document.
            String[] fieldNames = doc.MailMerge.GetFieldNames();
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

        [Ignore("WORDSNET-17733")]
        [Test]
        [TestCase("!", false, "")]
        [TestCase(", ", false, "")]
        [TestCase(" . ", false, "")]
        [TestCase(" :", false, "")]
        [TestCase("  ; ", false, "")]
        [TestCase(" ?  ", false, "")]
        [TestCase("  ¡  ", false, "")]
        [TestCase("  ¿  ", false, "")]
        [TestCase("!", true, "!\f")]
        [TestCase(", ", true, ", \f")]
        [TestCase(" . ", true, " . \f")]
        [TestCase(" :", true, " :\f")]
        [TestCase("  ; ", true, "  ; \f")]
        [TestCase(" ?  ", true, " ?  \f")]
        [TestCase("  ¡  ", true, "  ¡  \f")]
        [TestCase("  ¿  ", true, "  ¿  \f")]
        public void RemoveColonBetweenEmptyMergeFields(string punctuationMark,
            bool isCleanupParagraphsWithPunctuationMarks, string resultText)
        {
            //ExStart
            //ExFor:MailMerge.CleanupParagraphsWithPunctuationMarks
            //ExSummary:Shows how to remove paragraphs with punctuation marks after mail merge operation.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldMergeField mergeFieldOption1 = (FieldMergeField) builder.InsertField("MERGEFIELD", "Option_1");
            mergeFieldOption1.FieldName = "Option_1";

            // Here is the complete list of cleanable punctuation marks:
            // !
            // ,
            // .
            // :
            // ;
            // ?
            // ¡
            // ¿
            builder.Write(punctuationMark);

            FieldMergeField mergeFieldOption2 = (FieldMergeField) builder.InsertField("MERGEFIELD", "Option_2");
            mergeFieldOption2.FieldName = "Option_2";

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            // The default value of the option is true which means that the behaviour was changed to mimic MS Word
            // If you rely on the old behavior are able to revert it by setting the option to false
            doc.MailMerge.CleanupParagraphsWithPunctuationMarks = isCleanupParagraphsWithPunctuationMarks;

            doc.MailMerge.Execute(new[] { "Option_1", "Option_2" }, new object[] { null, null });

            doc.Save(ArtifactsDir + "RemoveColonBetweenEmptyMergeFields.docx");
            //ExEnd

            Assert.AreEqual(resultText, doc.GetText());
        }

        [Test]
        public void GetFieldNames()
        {
            //ExStart
            //ExFor:FieldAddressBlock
            //ExFor:FieldAddressBlock.GetFieldNames
            //ExSummary:Shows how to get mail merge field names used by the field
            Document doc = new Document(MyDir + "MailMerge.GetFieldNames.docx");

            string[] addressFieldsExpect =
            {
                "Company", "First Name", "Middle Name", "Last Name", "Suffix", "Address 1", "City", "State",
                "Country or Region", "Postal Code"
            };

            FieldAddressBlock addressBlockField = (FieldAddressBlock) doc.Range.Fields[0];
            string[] addressBlockFieldNames = addressBlockField.GetFieldNames();
            //ExEnd

            Assert.AreEqual(addressFieldsExpect, addressBlockFieldNames);

            string[] greetingFieldsExpect = { "Courtesy Title", "Last Name" };

            FieldGreetingLine greetingLineField = (FieldGreetingLine) doc.Range.Fields[1];
            string[] greetingLineFieldNames = greetingLineField.GetFieldNames();

            Assert.AreEqual(greetingFieldsExpect, greetingLineFieldNames);
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
        [TestCase(false,
            "\u0013MERGEFIELD \"testfield1\"\u0014«testfield1»\u0015value 1\u0013MERGEFIELD \"testfield3\"\u0014«testfield3»\u0015\f")]
        public void MustasheTemplateSyntax(bool restoreTags, String sectionText)
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
            table.Rows.Add("value 1");

            doc.MailMerge.Execute(table);

            String paraText = DocumentHelper.GetParagraphText(doc, 0);

            Assert.AreEqual(sectionText, paraText);
        }

        [Test]
        public void TestMailMergeGetRegionsHierarchy()
        {
            //ExStart
            //ExFor:MailMerge.GetRegionsHierarchy
            //ExFor:MailMergeRegionInfo
            //ExFor:MailMergeRegionInfo.Regions
            //ExFor:MailMergeRegionInfo.Name
            //ExFor:MailMergeRegionInfo.Fields
            //ExFor:MailMergeRegionInfo.StartField
            //ExFor:MailMergeRegionInfo.EndField
            //ExFor:MailMergeRegionInfo.Level
            //ExSummary:Shows how to get MailMergeRegionInfo and work with it
            Document doc = new Document(MyDir + "MailMerge.TestRegionsHierarchy.doc");

            //Returns a full hierarchy of regions (with fields) available in the document.
            MailMergeRegionInfo regionInfo = doc.MailMerge.GetRegionsHierarchy();

            //Get top regions in the document
            IList<MailMergeRegionInfo> topRegions = regionInfo.Regions;
            Assert.AreEqual(2, topRegions.Count);
            Assert.AreEqual("Region1", topRegions[0].Name);
            Assert.AreEqual("Region2", topRegions[1].Name);
            Assert.AreEqual(1, topRegions[0].Level);
            Assert.AreEqual(1, topRegions[1].Level);

            //Get nested region in first top region
            IList<MailMergeRegionInfo> nestedRegions = topRegions[0].Regions;
            Assert.AreEqual(2, nestedRegions.Count);
            Assert.AreEqual("NestedRegion1", nestedRegions[0].Name);
            Assert.AreEqual("NestedRegion2", nestedRegions[1].Name);
            Assert.AreEqual(2, nestedRegions[0].Level);
            Assert.AreEqual(2, nestedRegions[1].Level);

            //Get field list in first top region
            IList<Field> fieldList = topRegions[0].Fields;
            Assert.AreEqual(4, fieldList.Count);

            FieldMergeField startFieldMergeField = nestedRegions[0].StartField;
            Assert.AreEqual("TableStart:NestedRegion1", startFieldMergeField.FieldName);

            FieldMergeField endFieldMergeField = nestedRegions[0].EndField;
            Assert.AreEqual("TableEnd:NestedRegion1", endFieldMergeField.FieldName);
            //ExEnd
        }

        [Test]
        public void TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption()
        {
            //ExStart
            //ExFor:MailMerge.MailMergeCallback
            //ExFor:IMailMergeCallback
            //ExFor:IMailMergeCallback.TagsReplaced
            //ExSummary:Shows how to define custom logic for handling events during mail merge.
            Document document = new Document();
            document.MailMerge.UseNonMergeFields = true;

            MailMergeCallbackStub mailMergeCallbackStub = new MailMergeCallbackStub();
            document.MailMerge.MailMergeCallback = mailMergeCallbackStub;

            document.MailMerge.Execute(new String[0], new object[0]);

            Assert.AreEqual(1, mailMergeCallbackStub.TagsReplacedCounter);
        }

        private class MailMergeCallbackStub : IMailMergeCallback
        {
            public void TagsReplaced()
            {
                TagsReplacedCounter++;
            }

            public int TagsReplacedCounter { get; private set; }
        }
        //ExEnd

        [Test]
        [TestCase("Region1")]
        [TestCase("NestedRegion1")]
        public void GetRegionsByName(string regionName)
        {
            Document doc = new Document(MyDir + "MailMerge.RegionsByName.doc");

            IList<MailMergeRegionInfo> regions = doc.MailMerge.GetRegionsByName(regionName);
            Assert.AreEqual(2, regions.Count);

            foreach (MailMergeRegionInfo region in regions)
            {
                Assert.AreEqual(regionName, region.Name);
            }
        }

        [Test]
        public void CleanupOptions()
        {
            Document doc = new Document(MyDir + "MailMerge.CleanUp.docx");

            DataTable data = GetDataTable();

            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyTableRows;
            doc.MailMerge.ExecuteWithRegions(data);

            doc.Save(ArtifactsDir + "MailMerge.CleanUp.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "MailMerge.CleanUp.docx", GoldsDir + "MailMerge.CleanUp Gold.docx"));
        }

        /// <summary>
        /// Create DataTable and fill it with data.
        /// In real life this DataTable should be filled from a database.
        /// </summary>
        private static DataTable GetDataTable()
        {
            DataTable dataTable = new DataTable("StudentCourse");
            dataTable.Columns.Add("CourseName");

            DataRow dataRowEmpty = dataTable.NewRow();
            dataTable.Rows.Add(dataRowEmpty);
            dataRowEmpty[0] = string.Empty;

            for (int i = 0; i < 10; i++)
            {
                DataRow datarow = dataTable.NewRow();
                dataTable.Rows.Add(datarow);
                datarow[0] = "Course " + i;
            }

            return dataTable;
        }

        [Test] 
        public void UnconditionalMergeFieldsAndRegions()
        {
            //ExStart
            //ExFor:MailMerge.UnconditionalMergeFieldsAndRegions
            //ExSummary:Shows how to merge fields or regions regardless of the parent IF field's condition.
            Document doc = new Document(MyDir + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");

            // Merge fields and merge regions are merged regardless of the parent IF field's condition.
            doc.MailMerge.UnconditionalMergeFieldsAndRegions = true;

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                new string[] { "FullName" },
                new object[] { "James Bond" });

            doc.Save(ArtifactsDir + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");
            //ExEnd
        }
    }
}