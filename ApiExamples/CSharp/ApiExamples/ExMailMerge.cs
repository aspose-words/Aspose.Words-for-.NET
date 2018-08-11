// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words.Fields;
using Aspose.Words;
using Aspose.Words.MailMerging;
using NUnit.Framework;
#if !(NETSTANDARD2_0 || __MOBILE__)
using System.Web;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExMailMerge : ApiExampleBase
    {
#if !(NETSTANDARD2_0 || __MOBILE__)
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
            doc.MailMerge.Execute(new String[] { "FullName", "Company", "Address", "Address2", "City" }, new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
            Assert.That(() => doc.Save(Response, @"\Artifacts\MailMerge.ExecuteArray.doc", ContentDisposition.Inline, null), Throws.TypeOf<ArgumentNullException>()); //Thrown because HttpResponse is null in the test.
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

            doc.Save(MyDir + @"\Artifacts\MailMerge.ExecuteDataTable.doc");
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

        [Test]
        public void GetFieldNames()
        {
            //ExStart
            //ExFor:FieldAddressBlock.#ctor
            //ExFor:FieldAddressBlock.GetFieldNames
            //ExSummary:Shows how to get mail merge field names used by the field
            Document doc = new Document(MyDir + "MailMerge.GetFieldNames.docx");

            string[] addressFieldsExpect = { "Company", "First Name", "Middle Name", "Last Name", "Suffix", "Address 1", "City", "State", "Country or Region", "Postal Code" };

            FieldAddressBlock addressBlockField = (FieldAddressBlock)doc.Range.Fields[0]; 
            string[] addressBlockFieldNames = addressBlockField.GetFieldNames();
            //ExEnd
                                                                                         
            Assert.AreEqual(addressFieldsExpect, addressBlockFieldNames);

            string[] greetingFieldsExpect = { "Courtesy Title", "Last Name" };

            FieldGreetingLine greetingLineField = (FieldGreetingLine)doc.Range.Fields[1];
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
        [TestCase(false, "\u0013MERGEFIELD \"testfield1\"\u0014«testfield1»\u0015value 1\u0013MERGEFIELD \"testfield3\"\u0014«testfield3»\u0015\f")]
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
            Assert.AreEqual(topRegions[0].Name, "Region1");
            Assert.AreEqual(topRegions[1].Name, "Region2");
            Assert.AreEqual(1, topRegions[0].Level);
            Assert.AreEqual(1, topRegions[1].Level);

            //Get nested region in first top region
            IList<MailMergeRegionInfo> nestedRegions = topRegions[0].Regions;
            Assert.AreEqual(2, nestedRegions.Count);
            Assert.AreEqual(nestedRegions[0].Name, "NestedRegion1");
            Assert.AreEqual(nestedRegions[1].Name, "NestedRegion2");
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
            //ExFor:IMailMergeCallback
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
                mTagsReplacedCounter++;
            }

            public int TagsReplacedCounter
            {
                get { return mTagsReplacedCounter; }
            }

            private int mTagsReplacedCounter;
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

            doc.Save(MyDir + @"\Artifacts\MailMerge.CleanUp.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\MailMerge.CleanUp.docx", MyDir + @"\Golds\MailMerge.CleanUp Gold.docx"));
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
    }
}
