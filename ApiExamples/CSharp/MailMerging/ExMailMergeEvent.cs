// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;

using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.MailMerging;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExMailMergeEvent : ApiExampleBase
    {
        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void MailMergeInsertHtmlCaller()
        {
            this.MailMergeInsertHtml();
        }

        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(string)
        //ExFor:MailMerge.FieldMergingCallback
        //ExFor:IFieldMergingCallback
        //ExFor:FieldMergingArgs
        //ExFor:FieldMergingArgsBase.DocumentFieldName
        //ExFor:FieldMergingArgsBase.Document
        //ExFor:FieldMergingArgsBase.FieldValue
        //ExFor:IFieldMergingCallback.FieldMerging
        //ExFor:FieldMergingArgs.Text
        //ExSummary:Shows how to mail merge HTML data into a document.
        // File 'MailMerge.InsertHtml.doc' has merge field named 'htmlField1' in it.
        // File 'MailMerge.HtmlData.html' contains some valid Html data.
        // The same approach can be used when merging HTML data from database.
        public void MailMergeInsertHtml()
        {
            Document doc = new Document(MyDir + "MailMerge.InsertHtml.doc");

            // Add a handler for the MergeField event.
            doc.MailMerge.FieldMergingCallback = new HandleMergeFieldInsertHtml();

            // Load some Html from file.
            StreamReader sr = File.OpenText(MyDir + "MailMerge.HtmlData.html");
            string htmltext = sr.ReadToEnd();
            sr.Close();

            // Execute mail merge.
            doc.MailMerge.Execute(new string[] { "htmlField1" }, new string[] { htmltext });

            // Save resulting document with a new name.
            doc.Save(MyDir + @"\Artifacts\MailMerge.InsertHtml.doc");
        }

        private class HandleMergeFieldInsertHtml : IFieldMergingCallback
        {
            /// <summary>
            /// This is called when merge field is actually merged with data in the document.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                // All merge fields that expect HTML data should be marked with some prefix, e.g. 'html'.
                if (e.DocumentFieldName.StartsWith("html"))
                {
                    // Insert the text for this merge field as HTML data, using DocumentBuilder.
                    DocumentBuilder builder = new DocumentBuilder(e.Document);
                    builder.MoveToMergeField(e.DocumentFieldName);
                    builder.InsertHtml((string)e.FieldValue);

                    // The HTML text itself should not be inserted.
                    // We have already inserted it as an HTML.
                    e.Text = "";
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                // Do nothing.
            }
        }
        //ExEnd


        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void MailMergeInsertCheckBoxCaller()
        {
            this.MailMergeInsertCheckBox();
        }

        //ExStart
        //ExFor:DocumentBuilder.MoveToMergeField(string)
        //ExFor:DocumentBuilder.InsertCheckBox
        //ExFor:FieldMergingArgsBase.FieldName
        //ExSummary:Shows how to insert checkbox form fields into a document during mail merge.
        // File 'MailMerge.InsertCheckBox.doc' is a template
        // containing the table with the following fields in it:
        // <<TableStart:StudentCourse>> <<CourseName>> <<TableEnd:StudentCourse>>.
        public void MailMergeInsertCheckBox()
        {
            Document doc = new Document(MyDir + "MailMerge.InsertCheckBox.doc");

            // Add a handler for the MergeField event.
            doc.MailMerge.FieldMergingCallback = new HandleMergeFieldInsertCheckBox();

            // Execute mail merge with regions.
            DataTable dataTable = GetStudentCourseDataTable();
            doc.MailMerge.ExecuteWithRegions(dataTable);

            // Save resulting document with a new name.
            doc.Save(MyDir + @"\Artifacts\MailMerge.InsertCheckBox.doc");
        }

        private class HandleMergeFieldInsertCheckBox : IFieldMergingCallback
        {
            /// <summary>
            /// This is called for each merge field in the document
            /// when Document.MailMerge.ExecuteWithRegions is called.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (e.DocumentFieldName.Equals("CourseName"))
                {
                    // Insert the checkbox for this merge field, using DocumentBuilder.
                    DocumentBuilder builder = new DocumentBuilder(e.Document);
                    builder.MoveToMergeField(e.FieldName);
                    builder.InsertCheckBox(e.DocumentFieldName + this.mCheckBoxCount.ToString(), false, 0);
                    builder.Write((string)e.FieldValue);
                    this.mCheckBoxCount++;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }

            /// <summary>
            /// Counter for CheckBox name generation
            /// </summary>
            private int mCheckBoxCount;
        }

        /// <summary>
        /// Create DataTable and fill it with data.
        /// In real life this DataTable should be filled from a database.
        /// </summary>
        private static DataTable GetStudentCourseDataTable()
        {
            DataTable dataTable = new DataTable("StudentCourse");
            dataTable.Columns.Add("CourseName");
            for (int i = 0; i < 10; i++)
            {
                DataRow datarow = dataTable.NewRow();
                dataTable.Rows.Add(datarow);
                datarow[0] = "Course " + i.ToString();
            }
            return dataTable;
        }
        //ExEnd

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void MailMergeAlternatingRowsCaller()
        {
            this.MailMergeAlternatingRows();
        }

        //ExStart
        //ExId:MailMergeAlternatingRows
        //ExSummary:Demonstrates how to implement custom logic in the MergeField event to apply cell formatting.
        public void MailMergeAlternatingRows()
        {
            Document doc = new Document(MyDir + "MailMerge.AlternatingRows.doc");

            // Add a handler for the MergeField event.
            doc.MailMerge.FieldMergingCallback = new HandleMergeFieldAlternatingRows();

            // Execute mail merge with regions.
            DataTable dataTable = GetSuppliersDataTable();
            doc.MailMerge.ExecuteWithRegions(dataTable);

            doc.Save(MyDir + @"\Artifacts\MailMerge.AlternatingRows.doc");
        }

        private class HandleMergeFieldAlternatingRows : IFieldMergingCallback
        {
            /// <summary>
            /// Called for every merge field encountered in the document.
            /// We can either return some data to the mail merge engine or do something
            /// else with the document. In this case we modify cell formatting.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (this.mBuilder == null)
                    this.mBuilder = new DocumentBuilder(e.Document);

                // This way we catch the beginning of a new row.
                if (e.FieldName.Equals("CompanyName"))
                {
                    // Select the color depending on whether the row number is even or odd.
                    Color rowColor;
                    if (IsOdd(this.mRowIdx))
                        rowColor = Color.FromArgb(213, 227, 235);
                    else
                        rowColor = Color.FromArgb(242, 242, 242);

                    // There is no way to set cell properties for the whole row at the moment,
                    // so we have to iterate over all cells in the row.
                    for (int colIdx = 0; colIdx < 4; colIdx++)
                    {
                        this.mBuilder.MoveToCell(0, this.mRowIdx, colIdx, 0);
                        this.mBuilder.CellFormat.Shading.BackgroundPatternColor = rowColor;
                    }

                    this.mRowIdx++;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }

            private DocumentBuilder mBuilder;
            private int mRowIdx;
        }

        /// <summary>
        /// Returns true if the value is odd; false if the value is even.
        /// </summary>
        private static bool IsOdd(int value)
        {
            // The code is a bit complex, but otherwise automatic conversion to VB does not work.
            return ((value / 2) * 2).Equals(value);
        }

        /// <summary>
        /// Create DataTable and fill it with data.
        /// In real life this DataTable should be filled from a database.
        /// </summary>
        private static DataTable GetSuppliersDataTable()
        {
            DataTable dataTable = new DataTable("Suppliers");
            dataTable.Columns.Add("CompanyName");
            dataTable.Columns.Add("ContactName");
            for (int i = 0; i < 10; i++)
            {
                DataRow datarow = dataTable.NewRow();
                dataTable.Rows.Add(datarow);
                datarow[0] = "Company " + i.ToString();
                datarow[1] = "Contact " + i.ToString();
            }
            return dataTable;
        }
        //ExEnd

        [Test]
        public void MailMergeImageFromUrl()
        {
            //ExStart
            //ExFor:MailMerge.Execute(String[], Object[])
            //ExSummary:Demonstrates how to merge an image from a web address using an Image field.
            Document doc = new Document(MyDir + "MailMerge.MergeImageSimple.doc");

            // Pass a URL which points to the image to merge into the document.
            doc.MailMerge.Execute(new string[] { "Logo" }, new object[] { "http://www.aspose.com/images/aspose-logo.gif" });

            doc.Save(MyDir + @"\Artifacts\MailMerge.MergeImageFromUrl.doc");
            //ExEnd

            // Verify the image was merged into the document.
            Shape logoImage = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.IsNotNull(logoImage);
            Assert.IsTrue(logoImage.HasImage);
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void MailMergeImageFromBlobCaller()
        {
            this.MailMergeImageFromBlob();
        }
        
        //ExStart
        //ExFor:MailMerge.FieldMergingCallback
        //ExFor:MailMerge.ExecuteWithRegions(IDataReader,string)
        //ExFor:IFieldMergingCallback
        //ExFor:ImageFieldMergingArgs
        //ExFor:IFieldMergingCallback.FieldMerging
        //ExFor:IFieldMergingCallback.ImageFieldMerging
        //ExFor:ImageFieldMergingArgs.ImageStream
        //ExId:MailMergeImageFromBlob
        //ExSummary:Shows how to insert images stored in a database BLOB field into a report.
        public void MailMergeImageFromBlob()
        {
            Document doc = new Document(MyDir + "MailMerge.MergeImage.doc");

            // Set up the event handler for image fields.
            doc.MailMerge.FieldMergingCallback = new HandleMergeImageFieldFromBlob();

            // Open a database connection.
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DatabaseDir + "Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            // Open the data reader. It needs to be in the normal mode that reads all record at once.
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Employees", conn);
            IDataReader dataReader = cmd.ExecuteReader();

            // Perform mail merge.
            doc.MailMerge.ExecuteWithRegions(dataReader, "Employees");

            // Close the database.
            conn.Close();

            doc.Save(MyDir + @"\Artifacts\MailMerge.MergeImage.doc");
        }

        private class HandleMergeImageFieldFromBlob : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                // Do nothing.
            }

            /// <summary>
            /// This is called when mail merge engine encounters Image:XXX merge field in the document.
            /// You have a chance to return an Image object, file name or a stream that contains the image.
            /// </summary>
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                // The field value is a byte array, just cast it and create a stream on it.
                MemoryStream imageStream = new MemoryStream((byte[])e.FieldValue);
                // Now the mail merge engine will retrieve the image from the stream.
                e.ImageStream = imageStream;
            }
        }
        //ExEnd
    }
}
