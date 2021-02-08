// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String)
        //ExFor:MailMerge.FieldMergingCallback
        //ExFor:IFieldMergingCallback
        //ExFor:FieldMergingArgs
        //ExFor:FieldMergingArgsBase
        //ExFor:FieldMergingArgsBase.Field
        //ExFor:FieldMergingArgsBase.DocumentFieldName
        //ExFor:FieldMergingArgsBase.Document
        //ExFor:IFieldMergingCallback.FieldMerging
        //ExFor:FieldMergingArgs.Text
        //ExSummary:Shows how to execute a mail merge with a custom callback that handles merge data in the form of HTML documents.
        [Test] //ExSkip
        public void MergeHtml()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(@"MERGEFIELD  html_Title  \b Content");
            builder.InsertField(@"MERGEFIELD  html_Body  \b Content");

            object[] mergeData =
            {
                "<html>" +
                    "<h1>" +
                        "<span style=\"color: #0000ff; font-family: Arial;\">Hello World!</span>" +
                    "</h1>" +
                "</html>", 

                "<html>" +
                    "<blockquote>" +
                        "<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>" +
                    "</blockquote>" +
                "</html>"
            };
            
            doc.MailMerge.FieldMergingCallback = new HandleMergeFieldInsertHtml();
            doc.MailMerge.Execute(new[] { "html_Title", "html_Body" }, mergeData);
            
            doc.Save(ArtifactsDir + "MailMergeEvent.MergeHtml.docx");
        }

        /// <summary>
        /// If the mail merge encounters a MERGEFIELD whose name starts with the "html_" prefix,
        /// this callback parses its merge data as HTML content and adds the result to the document location of the MERGEFIELD.
        /// </summary>
        private class HandleMergeFieldInsertHtml : IFieldMergingCallback
        {
            /// <summary>
            /// Called when a mail merge merges data into a MERGEFIELD.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                if (args.DocumentFieldName.StartsWith("html_") && args.Field.GetFieldCode().Contains("\\b"))
                {
                    // Add parsed HTML data to the document's body.
                    DocumentBuilder builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.DocumentFieldName);
                    builder.InsertHtml((string)args.FieldValue);

                    // Since we have already inserted the merged content manually,
                    // we will not need to respond to this event by returning content via the "Text" property. 
                    args.Text = string.Empty;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }
        }
        //ExEnd

        //ExStart
        //ExFor:FieldMergingArgsBase.FieldValue
        //ExSummary:Shows how to edit values that MERGEFIELDs receive as a mail merge takes place.
        [Test] //ExSkip
        public void FieldFormats()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some MERGEFIELDs with format switches that will edit the values they will receive during a mail merge.
            builder.InsertField("MERGEFIELD text_Field1 \\* Caps", null);
            builder.Write(", ");
            builder.InsertField("MERGEFIELD text_Field2 \\* Upper", null);
            builder.Write(", ");
            builder.InsertField("MERGEFIELD numeric_Field1 \\# 0.0", null);

            builder.Document.MailMerge.FieldMergingCallback = new FieldValueMergingCallback();

            builder.Document.MailMerge.Execute(
                new string[] { "text_Field1", "text_Field2", "numeric_Field1" },
                new object[] { "Field 1", "Field 2", 10 });
            string t = doc.GetText().Trim();
            Assert.AreEqual("Merge Value For \"Text_Field1\": Field 1, MERGE VALUE FOR \"TEXT_FIELD2\": FIELD 2, 10000.0", doc.GetText().Trim());
        }

        /// <summary>
        /// Edits the values that MERGEFIELDs receive during a mail merge.
        /// The name of a MERGEFIELD must have a prefix for this callback to take effect on its value.
        /// </summary>
        private class FieldValueMergingCallback : IFieldMergingCallback
        {
            /// <summary>
            /// Called when a mail merge merges data into a MERGEFIELD.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (e.FieldName.StartsWith("text_"))
                    e.FieldValue = $"Merge value for \"{e.FieldName}\": {(string)e.FieldValue}";
                else if (e.FieldName.StartsWith("numeric_"))
                    e.FieldValue = (int)e.FieldValue * 1000;
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                // Do nothing.
            }
        }
        //ExEnd

        //ExStart
        //ExFor:DocumentBuilder.MoveToMergeField(String)
        //ExFor:FieldMergingArgsBase.FieldName
        //ExFor:FieldMergingArgsBase.TableName
        //ExFor:FieldMergingArgsBase.RecordIndex
        //ExSummary:Shows how to insert checkbox form fields into MERGEFIELDs as merge data during mail merge.
        [Test] //ExSkip
        public void InsertCheckBox()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use MERGEFIELDs with "TableStart"/"TableEnd" tags to define a mail merge region
            // which belongs to a data source named "StudentCourse" and has a MERGEFIELD which accepts data from a column named "CourseName".
            builder.StartTable();
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD  TableStart:StudentCourse ");
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD  CourseName ");
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD  TableEnd:StudentCourse ");
            builder.EndTable();

            doc.MailMerge.FieldMergingCallback = new HandleMergeFieldInsertCheckBox();

            DataTable dataTable = GetStudentCourseDataTable();

            doc.MailMerge.ExecuteWithRegions(dataTable);
            doc.Save(ArtifactsDir + "MailMergeEvent.InsertCheckBox.docx");
            TestUtil.MailMergeMatchesDataTable(dataTable, new Document(ArtifactsDir + "MailMergeEvent.InsertCheckBox.docx"), false); //ExSkip
        }

        /// <summary>
        /// Upon encountering a MERGEFIELD with a specific name, inserts a check box form field instead of merge data text.
        /// </summary>
        private class HandleMergeFieldInsertCheckBox : IFieldMergingCallback
        {
            /// <summary>
            /// Called when a mail merge merges data into a MERGEFIELD.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                if (args.DocumentFieldName.Equals("CourseName"))
                {
                    Assert.AreEqual("StudentCourse", args.TableName);

                    DocumentBuilder builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.FieldName);
                    builder.InsertCheckBox(args.DocumentFieldName + mCheckBoxCount, false, 0);

                    string fieldValue = args.FieldValue.ToString();

                    // In this case, for every record index 'n', the corresponding field value is "Course n".
                    Assert.AreEqual(char.GetNumericValue(fieldValue[7]), args.RecordIndex);

                    builder.Write(fieldValue);
                    mCheckBoxCount++;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }

            private int mCheckBoxCount;
        }

        /// <summary>
        /// Creates a mail merge data source.
        /// </summary>
        private static DataTable GetStudentCourseDataTable()
        {
            DataTable dataTable = new DataTable("StudentCourse");
            dataTable.Columns.Add("CourseName");
            for (int i = 0; i < 10; i++)
            {
                DataRow datarow = dataTable.NewRow();
                dataTable.Rows.Add(datarow);
                datarow[0] = "Course " + i;
            }

            return dataTable;
        }
        //ExEnd

        //ExStart
        //ExFor:MailMerge.ExecuteWithRegions(DataTable)
        //ExSummary:Demonstrates how to format cells during a mail merge.
        [Test] //ExSkip
        public void AlternatingRows()
        {
            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");

            doc.MailMerge.FieldMergingCallback = new HandleMergeFieldAlternatingRows();

            DataTable dataTable = GetSuppliersDataTable();
            doc.MailMerge.ExecuteWithRegions(dataTable);

            doc.Save(ArtifactsDir + "MailMergeEvent.AlternatingRows.docx");
            TestUtil.MailMergeMatchesDataTable(dataTable, new Document(ArtifactsDir + "MailMergeEvent.AlternatingRows.docx"), false); //ExSkip
        }

        /// <summary>
        /// Formats table rows as a mail merge takes place to alternate between two colors on odd/even rows.
        /// </summary>
        private class HandleMergeFieldAlternatingRows : IFieldMergingCallback
        {
            /// <summary>
            /// Called when a mail merge merges data into a MERGEFIELD.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                if (mBuilder == null)
                    mBuilder = new DocumentBuilder(args.Document);

                // This is true of we are on the first column, which means we have moved to a new row.
                if (args.FieldName.Equals("CompanyName"))
                {
                    Color rowColor = IsOdd(mRowIdx) ? Color.FromArgb(213, 227, 235) : Color.FromArgb(242, 242, 242);

                    for (int colIdx = 0; colIdx < 4; colIdx++)
                    {
                        mBuilder.MoveToCell(0, mRowIdx, colIdx, 0);
                        mBuilder.CellFormat.Shading.BackgroundPatternColor = rowColor;
                    }

                    mRowIdx++;
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
        /// Function needed for Visual Basic autoporting that returns the parity of the passed number.
        /// </summary>
        private static bool IsOdd(int value)
        {
            return (value / 2 * 2).Equals(value);
        }

        /// <summary>
        /// Creates a mail merge data source.
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
                datarow[0] = "Company " + i;
                datarow[1] = "Contact " + i;
            }

            return dataTable;
        }
        //ExEnd

        [Test]
        public void ImageFromUrl()
        {
            //ExStart
            //ExFor:MailMerge.Execute(String[], Object[])
            //ExSummary:Shows how to merge an image from a URI as mail merge data into a MERGEFIELD.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // MERGEFIELDs with "Image:" tags will receive an image during a mail merge.
            // The string after the colon in the "Image:" tag corresponds to a column name
            // in the data source whose cells contain URIs of image files.
            builder.InsertField("MERGEFIELD  Image:logo_FromWeb ");
            builder.InsertField("MERGEFIELD  Image:logo_FromFileSystem ");

            // Create a data source that contains URIs of images that we will merge. 
            // A URI can be a web URL that points to an image, or a local file system filename of an image file.
            string[] columns = { "logo_FromWeb", "logo_FromFileSystem" };
            object[] URIs = { AsposeLogoUrl, ImageDir + "Logo.jpg" };

            // Execute a mail merge on a data source with one row.
            doc.MailMerge.Execute(columns, URIs);

            doc.Save(ArtifactsDir + "MailMergeEvent.ImageFromUrl.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "MailMergeEvent.ImageFromUrl.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(320, 320, ImageType.Png, imageShape);
        }

#if NET462 || JAVA
        //ExStart
        //ExFor:MailMerge.FieldMergingCallback
        //ExFor:MailMerge.ExecuteWithRegions(IDataReader,String)
        //ExFor:IFieldMergingCallback
        //ExFor:ImageFieldMergingArgs
        //ExFor:IFieldMergingCallback.FieldMerging
        //ExFor:IFieldMergingCallback.ImageFieldMerging
        //ExFor:ImageFieldMergingArgs.ImageStream
        //ExSummary:Shows how to insert images stored in a database BLOB field into a report.
        [Test, Category("SkipMono")] //ExSkip        
        public void ImageFromBlob()
        {
            Document doc = new Document(MyDir + "Mail merge destination - Northwind employees.docx");

            doc.MailMerge.FieldMergingCallback = new HandleMergeImageFieldFromBlob();

            string connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={DatabaseDir + "Northwind.mdb"};";
            string query = "SELECT FirstName, LastName, Title, Address, City, Region, Country, PhotoBLOB FROM Employees";

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                conn.Open();

                // Open the data reader, which needs to be in a mode that reads all records at once.
                OleDbCommand cmd = new OleDbCommand(query, conn);
                IDataReader dataReader = cmd.ExecuteReader();

                doc.MailMerge.ExecuteWithRegions(dataReader, "Employees");
            }

            doc.Save(ArtifactsDir + "MailMergeEvent.ImageFromBlob.docx");
            TestUtil.MailMergeMatchesQueryResult(DatabaseDir + "Northwind.mdb", query, new Document(ArtifactsDir + "MailMergeEvent.ImageFromBlob.docx"), false); //ExSkip
        }

        private class HandleMergeImageFieldFromBlob : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                // Do nothing.
            }

            /// <summary>
            /// This is called when a mail merge encounters a MERGEFIELD in the document with an "Image:" tag in its name.
            /// </summary>
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                MemoryStream imageStream = new MemoryStream((byte[])e.FieldValue);
                e.ImageStream = imageStream;
            }
        }
        //ExEnd
#endif
    }
}