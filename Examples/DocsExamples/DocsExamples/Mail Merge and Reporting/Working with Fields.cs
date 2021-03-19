using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace DocsExamples.Mail_Merge_and_Reporting
{
    internal class WorkingWithFields : DocsExamplesBase
    {
        [Test]
        public void MailMergeFormFields()
        {
            //ExStart:MailMergeFormFields
            Document doc = new Document(MyDir + "Mail merge destinations - Fax.docx");

            // Setup mail merge event handler to do the custom work.
            doc.MailMerge.FieldMergingCallback = new HandleMergeField();
            // Trim trailing and leading whitespaces mail merge values.
            doc.MailMerge.TrimWhitespaces = false;

            string[] fieldNames = {
                "RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
                "Subject", "Body", "Urgent", "ForReview", "PleaseComment"
            };

            object[] fieldValues = {
                "Josh", "Jenny", "123456789", "", "Hello",
                "<b>HTML Body Test message 1</b>", true, false, true
            };

            doc.MailMerge.Execute(fieldNames, fieldValues);

            doc.Save(ArtifactsDir + "WorkingWithFields.MailMergeFormFields.docx");
            //ExEnd:MailMergeFormFields
        }

        //ExStart:HandleMergeField
        private class HandleMergeField : IFieldMergingCallback
        {
            /// <summary>
            /// This handler is called for every mail merge field found in the document,
            /// for every record found in the data source.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (mBuilder == null)
                    mBuilder = new DocumentBuilder(e.Document);

                // We decided that we want all boolean values to be output as check box form fields.
                if (e.FieldValue is bool)
                {
                    // Move the "cursor" to the current merge field.
                    mBuilder.MoveToMergeField(e.FieldName);

                    string checkBoxName = $"{e.FieldName}{e.RecordIndex}";

                    mBuilder.InsertCheckBox(checkBoxName, (bool) e.FieldValue, 0);

                    return;
                }

                switch (e.FieldName)
                {
                    case "Body":
                        mBuilder.MoveToMergeField(e.FieldName);
                        mBuilder.InsertHtml((string) e.FieldValue);
                        break;
                    case "Subject":
                    {
                        mBuilder.MoveToMergeField(e.FieldName);
                        string textInputName = $"{e.FieldName}{e.RecordIndex}";
                        mBuilder.InsertTextInput(textInputName, TextFormFieldType.Regular, "", (string) e.FieldValue, 0);
                        break;
                    }
                }
            }

            //ExStart:ImageFieldMerging
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                args.ImageFileName = "Image.png";
                args.ImageWidth.Value = 200;
                args.ImageHeight = new MergeFieldImageDimension(200, MergeFieldImageDimensionUnit.Percent);
            }
            //ExEnd:ImageFieldMerging

            private DocumentBuilder mBuilder;
        }
        //ExEnd:HandleMergeField

        [Test]
        public void MailMergeImageField()
        {
            //ExStart:MailMergeImageField       
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("{{#foreach example}}");
            builder.Writeln("{{Image(126pt;126pt):stempel}}");
            builder.Writeln("{{/foreach example}}");

            doc.MailMerge.UseNonMergeFields = true;
            doc.MailMerge.TrimWhitespaces = true;
            doc.MailMerge.UseWholeParagraphAsRegion = false;
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyTableRows
                    | MailMergeCleanupOptions.RemoveContainingFields
                    | MailMergeCleanupOptions.RemoveUnusedRegions
                    | MailMergeCleanupOptions.RemoveUnusedFields;

            doc.MailMerge.FieldMergingCallback = new ImageFieldMergingHandler();
            doc.MailMerge.ExecuteWithRegions(new DataSourceRoot());

            doc.Save(ArtifactsDir + "WorkingWithFields.MailMergeImageField.docx");
            //ExEnd:MailMergeImageField
        }

        //ExStart:ImageFieldMergingHandler
        private class ImageFieldMergingHandler : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                //  Implementation is not required.
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                Shape shape = new Shape(args.Document, ShapeType.Image)
                {
                    Width = 126, Height = 126, WrapType = WrapType.Square
                };

                shape.ImageData.SetImage(MyDir + "Mail merge image.png");

                args.Shape = shape;
            }
        }
        //ExEnd:ImageFieldMergingHandler

        //ExStart:DataSourceRoot
        public class DataSourceRoot : IMailMergeDataSourceRoot
        {
            public IMailMergeDataSource GetDataSource(string s)
            {
                return new DataSource();
            }

            private class DataSource : IMailMergeDataSource
            {
                private bool next = true;

                string IMailMergeDataSource.TableName => TableName();

                private string TableName()
                {
                    return "example";
                }

                public bool MoveNext()
                {
                    bool result = next;
                    next = false;
                    return result;
                }

                public IMailMergeDataSource GetChildDataSource(string s)
                {
                    return null;
                }

                public bool GetValue(string fieldName, out object fieldValue)
                {
                    fieldValue = null;
                    return false;
                }
            }
        }
        //ExEnd:DataSourceRoot

        [Test]
        public void MailMergeAndConditionalField()
        {
            //ExStart:MailMergeAndConditionalField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEFIELD nested inside an IF field.
            // Since the IF field statement is false, the result of the inner MERGEFIELD will not be displayed,
            // and the MERGEFIELD will not receive any data during a mail merge.
            FieldIf fieldIf = (FieldIf)builder.InsertField(" IF 1 = 2 ");
            builder.MoveTo(fieldIf.Separator);
            builder.InsertField(" MERGEFIELD  FullName ");

            // We can still count MERGEFIELDs inside false-statement IF fields if we set this flag to true.
            doc.MailMerge.UnconditionalMergeFieldsAndRegions = true;

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FullName");
            dataTable.Rows.Add("James Bond");

            doc.MailMerge.Execute(dataTable);

            // The result will not be visible in the document because the IF field is false,
            // but the inner MERGEFIELD did indeed receive data.
            doc.Save(ArtifactsDir + "WorkingWithFields.MailMergeAndConditionalField.docx");
            //ExEnd:MailMergeAndConditionalField
        }

        [Test]
        public void MailMergeImageFromBlob()
        {
            //ExStart:MailMergeImageFromBlob
            Document doc = new Document(MyDir + "Mail merge destination - Northwind employees.docx");

            doc.MailMerge.FieldMergingCallback = new HandleMergeImageFieldFromBlob();

            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DatabaseDir + "Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Employees", conn);
            IDataReader dataReader = cmd.ExecuteReader();

            doc.MailMerge.ExecuteWithRegions(dataReader, "Employees");

            conn.Close();
            
            doc.Save(ArtifactsDir + "WorkingWithFields.MailMergeImageFromBlob.docx");
            //ExEnd:MailMergeImageFromBlob
        }

        //ExStart:HandleMergeImageFieldFromBlob 
        public class HandleMergeImageFieldFromBlob : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                // Do nothing.
            }

            /// <summary>
            /// This is called when mail merge engine encounters Image:XXX merge field in the document.
            /// You have a chance to return an Image object, file name, or a stream that contains the image.
            /// </summary>
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                // The field value is a byte array, just cast it and create a stream on it.
                MemoryStream imageStream = new MemoryStream((byte[]) e.FieldValue);
                // Now the mail merge engine will retrieve the image from the stream.
                e.ImageStream = imageStream;
            }
        }
        //ExEnd:HandleMergeImageFieldFromBlob

        [Test]
        public void HandleMailMergeSwitches()
        {
            Document doc = new Document(MyDir + "Field sample - MERGEFIELD.docx");

            doc.MailMerge.FieldMergingCallback = new MailMergeSwitches();

            const string html = @"<html>
                    <h1>Hello world!</h1>
            </html>";

            doc.MailMerge.Execute(new string[] { "htmlField1" }, new object[] { html });

            doc.Save(ArtifactsDir + "WorkingWithFields.HandleMailMergeSwitches.docx");
        }

        //ExStart:HandleMailMergeSwitches
        public class MailMergeSwitches : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (e.FieldName.StartsWith("HTML"))
                {
                    if (e.Field.GetFieldCode().Contains("\\b"))
                    {
                        FieldMergeField field = e.Field;

                        DocumentBuilder builder = new DocumentBuilder(e.Document);
                        builder.MoveToMergeField(e.DocumentFieldName, true, false);
                        builder.Write(field.TextBefore);
                        builder.InsertHtml(e.FieldValue.ToString());

                        e.Text = "";
                    }
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
            }
        }
        //ExEnd:HandleMailMergeSwitches

        [Test]
        public void AlternatingRows()
        {
            //ExStart:MailMergeAlternatingRows
            Document doc = new Document(MyDir + "Mail merge destination - Northwind suppliers.docx");

            doc.MailMerge.FieldMergingCallback = new HandleMergeFieldAlternatingRows();

            DataTable dataTable = GetSuppliersDataTable();
            doc.MailMerge.ExecuteWithRegions(dataTable);
            
            doc.Save(ArtifactsDir + "WorkingWithFields.AlternatingRows.doc");
            //ExEnd:MailMergeAlternatingRows
        }

        //ExStart:HandleMergeFieldAlternatingRows
        private class HandleMergeFieldAlternatingRows : IFieldMergingCallback
        {
            /// <summary>
            /// Called for every merge field encountered in the document.
            /// We can either return some data to the mail merge engine or do something else with the document.
            /// In this case we modify cell formatting.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (mBuilder == null)
                    mBuilder = new DocumentBuilder(e.Document);

                if (e.FieldName == "CompanyName")
                {
                    // Select the color depending on whether the row number is even or odd.
                    Color rowColor = IsOdd(mRowIdx) 
                        ? Color.FromArgb(213, 227, 235) 
                        : Color.FromArgb(242, 242, 242);

                    // There is no way to set cell properties for the whole row at the moment, so we have to iterate over all cells in the row.
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
        /// Returns true if the value is odd; false if the value is even.
        /// </summary>
        private static bool IsOdd(int value)
        {
            return (value / 2 * 2) == value;
        }

        /// <summary>
        /// Create DataTable and fill it with data.
        /// In real life this DataTable should be filled from a database.
        /// </summary>
        private DataTable GetSuppliersDataTable()
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
        //ExEnd:HandleMergeFieldAlternatingRows
    }
}