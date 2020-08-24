// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;
using LoadOptions = Aspose.Words.LoadOptions;
#if NET462 || JAVA
using Aspose.BarCode.BarCodeRecognition;
#elif NETCOREAPP2_1
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExField : ApiExampleBase
    {
        [Test]
        public void GetFieldFromDocument()
        {
            //ExStart
            //ExFor:FieldType
            //ExFor:FieldChar
            //ExFor:FieldChar.FieldType
            //ExFor:FieldChar.IsDirty
            //ExFor:FieldChar.IsLocked
            //ExFor:FieldChar.GetField
            //ExFor:Field.IsLocked
            //ExSummary:Shows how to work with a FieldStart node.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldDate field = (FieldDate)builder.InsertField(FieldType.FieldDate, true);
            field.Format.DateTimeFormat = "dddd, MMMM dd, yyyy";
            field.Update();
            
            FieldChar fieldStart = field.Start;

            Assert.AreEqual(FieldType.FieldDate, fieldStart.FieldType);
            Assert.AreEqual(false, fieldStart.IsDirty);
            Assert.AreEqual(false, fieldStart.IsLocked);

            // Retrieve the facade object which represents the field in the document.
            field = (FieldDate)fieldStart.GetField();

            Assert.AreEqual(false, field.IsLocked);
            Assert.AreEqual(" DATE  \\@ \"dddd, MMMM dd, yyyy\"", field.GetFieldCode());

            // Update the field to show the current date.
            field.Update();         
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            TestUtil.VerifyField(FieldType.FieldDate, " DATE  \\@ \"dddd, MMMM dd, yyyy\"", DateTime.Now.ToString("dddd, MMMM dd, yyyy"), doc.Range.Fields[0]);
        }
        
        [Test]
        public void GetFieldCode()
        {
            //ExStart
            //ExFor:Field.GetFieldCode
            //ExFor:Field.GetFieldCode(bool)
            //ExSummary:Shows how to get a field's field code.
            // Open a document which contains a MERGEFIELD inside an IF field.
            Document doc = new Document(MyDir + "Nested fields.docx");
            FieldIf fieldIf = (FieldIf)doc.Range.Fields[0];

            // There are two ways of getting a field's field code:
            // 1 -  Omit its inner fields.
            Assert.AreEqual(" IF  > 0 \" (surplus of ) \" \"\" ", fieldIf.GetFieldCode(false));

            // 2 -  Include its inner fields.
            Assert.AreEqual($" IF \u0013 MERGEFIELD NetIncome \u0014\u0015 > 0 \" (surplus of \u0013 MERGEFIELD  NetIncome \\f $ \u0014\u0015) \" \"\" ",
                fieldIf.GetFieldCode(true));

            // All inner nested fields are included by default.
            Assert.AreEqual(fieldIf.GetFieldCode(), fieldIf.GetFieldCode(true));
            //ExEnd
        }

        [Test]
        public void DisplayResult()
        {
            //ExStart
            //ExFor:Field.DisplayResult
            //ExSummary:Shows how to get the real text that a field displays in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("This document was written by ");
            FieldAuthor fieldAuthor = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);
            fieldAuthor.AuthorName = "John Doe";

            // We can use the DisplayResult attribute to verify what exact text
            // a field would display in its place in the document.
            Assert.AreEqual(string.Empty, fieldAuthor.DisplayResult);

            // Fields do not maintain accurate result values in real time. 
            // To make sure our fields display the accurate result at any given time,
            // such as right before a save operation, we need to manually update them.
            fieldAuthor.Update();

            Assert.AreEqual("John Doe", fieldAuthor.DisplayResult);

            doc.Save(ArtifactsDir + "Field.DisplayResult.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.DisplayResult.docx");

            Assert.AreEqual("John Doe", doc.Range.Fields[0].DisplayResult);
        }

        [Test]
        public void CreateWithFieldBuilder()
        {
            //ExStart
            //ExFor:FieldBuilder.#ctor(FieldType)
            //ExFor:FieldBuilder.BuildAndInsert(Inline)
            //ExSummary:Shows how to create and insert a field using a field builder.
            Document doc = new Document();

            // A convenient way of adding text content to a document is with a document builder.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write(" Hello world! This text is one Run, which is an inline node.");

            // Fields have their own builder, which we can use to construct a field code piece by piece.
            // In this case, we will construct a BARCODE field which represents a US postal code,
            // and then insert it in front of a Run.
            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldBarcode);
            fieldBuilder.AddArgument("90210");
            fieldBuilder.AddSwitch("\\f", "A");
            fieldBuilder.AddSwitch("\\u");

            fieldBuilder.BuildAndInsert(doc.FirstSection.Body.FirstParagraph.Runs[0]);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.CreateWithFieldBuilder.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.CreateWithFieldBuilder.docx");

            TestUtil.VerifyField(FieldType.FieldBarcode, " BARCODE 90210 \\f A \\u ", string.Empty, doc.Range.Fields[0]);

            Assert.AreEqual(doc.FirstSection.Body.FirstParagraph.Runs[11].PreviousSibling, doc.Range.Fields[0].End);
            Assert.AreEqual($"{ControlChar.FieldStartChar} BARCODE 90210 \\f A \\u {ControlChar.FieldEndChar} Hello world! This text is one Run, which is an inline node.", 
                doc.GetText().Trim());
        }

        [Test]
        public void RevNum()
        {
            //ExStart
            //ExFor:BuiltInDocumentProperties.RevisionNumber
            //ExFor:FieldRevNum
            //ExSummary:Shows how to work with REVNUM fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Current revision #");

            // Insert a REVNUM field, which displays the document's current revision number property.
            FieldRevNum field = (FieldRevNum)builder.InsertField(FieldType.FieldRevisionNum, true);

            Assert.AreEqual(" REVNUM ", field.GetFieldCode());
            Assert.AreEqual("1", field.Result);
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.RevisionNumber);

            // This property counts how many times a document has been saved in Microsoft Word,
            // and is unrelated to tracked revisions. We can find it by right clicking the document in Windows Explorer
            // via Properties -> Details. We can update this property manually.
            doc.BuiltInDocumentProperties.RevisionNumber++;
            Assert.AreEqual("1", field.Result); //ExSkip
            field.Update();

            Assert.AreEqual("2", field.Result);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            Assert.AreEqual(2, doc.BuiltInDocumentProperties.RevisionNumber);

            TestUtil.VerifyField(FieldType.FieldRevisionNum, " REVNUM ", "2", doc.Range.Fields[0]);
        }

        [Test]
        public void CreateInfoFieldWithFieldBuilder()
        {
            Document doc = new Document();
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 0);

            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldInfo);
            fieldBuilder.BuildAndInsert(run);

            doc.UpdateFields();
            doc = DocumentHelper.SaveOpen(doc);

            FieldInfo info = (FieldInfo)doc.Range.Fields[0];
            Assert.NotNull(info);
        }

        [Test]
        public void CreateInfoFieldWithDocumentBuilder()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("INFO MERGEFORMAT");

            doc = DocumentHelper.SaveOpen(doc);

            FieldInfo info = (FieldInfo)doc.Range.Fields[0];
            Assert.NotNull(info);
        }

        [Test]
        public void GetFieldFromFieldCollection()
        {
            Document doc = new Document(MyDir + "Table of contents.docx");

            Field field = doc.Range.Fields[0];

            // This should be the first field in the document - a TOC field
            Console.WriteLine(field.Type);
        }

        [Test]
        public void InsertFieldNone()
        {
            //ExStart
            //ExFor:FieldUnknown
            //ExSummary:Shows how to work with 'FieldNone' field in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a field that does not denote a real field type in its field code.
            Field field = builder.InsertField(" NOTAREALFIELD //a");

            // Fields like that can be created and read, and are assigned a special "FieldNone" type.
            Assert.AreEqual(FieldType.FieldNone, field.Type);

            // We can also still work with these fields, and assign them as instances of the FieldUnknown class.
            FieldUnknown fieldUnknown = (FieldUnknown)field;
            Assert.AreEqual(" NOTAREALFIELD //a", fieldUnknown.GetFieldCode());
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            TestUtil.VerifyField(FieldType.FieldNone, " NOTAREALFIELD //a", "Error! Bookmark not defined.", doc.Range.Fields[0]);
        }

        [Test]
        public void InsertTcField()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TC field at the current document builder position
            builder.InsertField("TC \"Entry Text\" \\f t");
        }

        [Test]
        public void FieldLocale()
        {
            //ExStart
            //ExFor:Field.LocaleId
            //ExSummary:Shows how to insert a field and work with its locale.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a DATE field, and then print the date it will display.
            // Your thread's current culture determines the formatting of the date.
            Field field = builder.InsertField(@"DATE");
            Console.WriteLine($"Today's date, as displayed in the \"{CultureInfo.CurrentCulture.EnglishName}\" culture: {field.Result}");

            Assert.AreEqual(1033, field.LocaleId);
            Assert.AreEqual(FieldUpdateCultureSource.CurrentThread, doc.FieldOptions.FieldUpdateCultureSource); //ExSkip

            // Changing the culture of our thread will impact the result of the DATE field.
            // Another way to get the DATE field to display a date in a different culture is to use the field's LocaleId attribute.
            // This way allows us to avoid changing the thread's culture to get this effect.
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            CultureInfo de = new CultureInfo("de-DE");
            field.LocaleId = de.LCID;
            field.Update();

            Console.WriteLine($"Today's date, as displayed according to the \"{CultureInfo.GetCultureInfo(field.LocaleId).EnglishName}\" culture: {field.Result}");
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            field = doc.Range.Fields[0]; 

            TestUtil.VerifyField(FieldType.FieldDate, "DATE", DateTime.Now.ToString(de.DateTimeFormat.ShortDatePattern), field);
            Assert.AreEqual(new CultureInfo("de-DE").LCID, field.LocaleId);
        }

        [Test]
        public void ChangeLocale()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("MERGEFIELD Date");

            // Store the current culture so it can be set back once mail merge is complete
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            // Set to German language so dates and numbers are formatted using this culture during mail merge
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

            doc.MailMerge.Execute(new[] { "Date" }, new object[] { DateTime.Now });

            // Restore the original culture and save the document
            Thread.CurrentThread.CurrentCulture = currentCulture;
            doc.Save(ArtifactsDir + "Field.ChangeLocale.docx");
        }

        [Test]
        public void RemoveTocFromDocument()
        {
            // Open a document which contains a TOC
            Document doc = new Document(MyDir + "Table of contents.docx");
            
            // Remove the first TOC from the document
            Field tocField = doc.Range.Fields[0];
            tocField.Remove();

            doc.Save(ArtifactsDir + "Field.RemoveTocFromDocument.docx");
        }

        [Test]
        public void InsertTcFieldsAtText()
        {
            Document doc = new Document();

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new InsertTcFieldHandler("Chapter 1", "\\l 1");

            // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document
            doc.Range.Replace(new Regex("The Beginning"), "", options);
        }

        private class InsertTcFieldHandler : IReplacingCallback
        {
            // Store the text and switches to be used for the TC fields
            private readonly string mFieldText;
            private readonly string mFieldSwitches;

            /// <summary>
            /// The display text and switches to use for each TC field. Display name can be an empty String or null.
            /// </summary>
            public InsertTcFieldHandler(string text, string switches)
            {
                mFieldText = text;
                mFieldSwitches = switches;
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                // Create a builder to insert the field
                DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
                // Move to the first node of the match
                builder.MoveTo(args.MatchNode);

                // If the user specified text to be used in the field as display text then use that, otherwise use the 
                // match String as the display text
                string insertText = !string.IsNullOrEmpty(mFieldText) ? mFieldText : args.Match.Value;

                // Insert the TC field before this node using the specified String as the display text and user defined switches
                builder.InsertField($"TC \"{insertText}\" {mFieldSwitches}");

                // We have done what we want so skip replacement
                return ReplaceAction.Skip;
            }
        }

        [TestCase(true)]
        [TestCase(false)]
        [Ignore("WORDSNET-16037")]
        public void UpdateDirtyFields(bool updateDirtyFields)
        {
            //ExStart
            //ExFor:Field.IsDirty
            //ExFor:LoadOptions.UpdateDirtyFields
            //ExSummary:Shows how to use special property for updating field result.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Give the document's built in "Author" property a value, and then display it with a field.
            doc.BuiltInDocumentProperties.Author = "John Doe";
            FieldAuthor field = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);

            Assert.False(field.IsDirty);
            Assert.AreEqual("John Doe", field.Result);

            // Update the property. The field still displays the old value.
            doc.BuiltInDocumentProperties.Author = "John & Jane Doe";

            Assert.AreEqual("John Doe", field.Result);

            // Since the field's value is out of date, we can mark it as "dirty".
            // This value will stay out of date until we update the field manually with the Field.Update() method.
            field.IsDirty = true;
            
            using (MemoryStream docStream = new MemoryStream())
            {
                // If we save without calling an update method,
                // the field will keep displaying the out of date value in the output document.
                doc.Save(docStream, SaveFormat.Docx);

                // The LoadOptions object has an option to update all fields
                // marked as "dirty" when loading the document.
                LoadOptions options = new LoadOptions();
                options.UpdateDirtyFields = updateDirtyFields;
                doc = new Document(docStream, options);
                
                Assert.AreEqual("John & Jane Doe", doc.BuiltInDocumentProperties.Author);

                field = (FieldAuthor)doc.Range.Fields[0];

                // Updating dirty fields like this automatically sets their "IsDirty" flag to false.
                if (updateDirtyFields)
                {
                    Assert.AreEqual("John & Jane Doe", field.Result);
                    Assert.False(field.IsDirty);
                }
                else
                {
                    Assert.AreEqual("John Doe", field.Result);
                    Assert.True(field.IsDirty);
                }
            }
            //ExEnd
        }

        [Test]
        public void InsertFieldWithFieldBuilderException()
        {
            Document doc = new Document();

            // Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!", 0);

            FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
            argumentBuilder.AddField(new FieldBuilder(FieldType.FieldMergeField));
            argumentBuilder.AddNode(run);
            argumentBuilder.AddText("Text argument builder");

            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldIncludeText);

            Assert.That(
                () => fieldBuilder.AddArgument(argumentBuilder).AddArgument("=").AddArgument("BestField")
                    .AddArgument(10).AddArgument(20.0).BuildAndInsert(run), Throws.TypeOf<ArgumentException>());
        }

#if NET462 || JAVA
        [Test]
        public void BarCodeWord2Pdf()
        {
            Document doc = new Document(MyDir + "Field sample - BARCODE.docx");

            // Set custom barcode generator
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            doc.Save(ArtifactsDir + "Field.BarCodeWord2Pdf.pdf");

            BarCodeReader barCode = BarCodeReaderPdf(ArtifactsDir + "Field.BarCodeWord2Pdf.pdf");
            Assert.AreEqual("QR", barCode.GetCodeType().ToString());
        }

        private BarCodeReader BarCodeReaderPdf(string filename)
        {
            // Set license for Aspose.BarCode
            Aspose.BarCode.License licenceBarCode = new Aspose.BarCode.License();
            licenceBarCode.SetLicense(LicenseDir + "Aspose.Total.NET.lic");

            // Bind the pdf document
            Aspose.Pdf.Facades.PdfExtractor pdfExtractor = new Aspose.Pdf.Facades.PdfExtractor();
            pdfExtractor.BindPdf(filename);

            // Set page range for image extraction
            pdfExtractor.StartPage = 1;
            pdfExtractor.EndPage = 1;

            pdfExtractor.ExtractImage();

            // Save image to stream
            MemoryStream imageStream = new MemoryStream();
            pdfExtractor.GetNextImage(imageStream);
            imageStream.Position = 0;

            // Recognize the barcode from the image stream above
            BarCodeReader barcodeReader = new BarCodeReader(imageStream, DecodeType.QR);
            while (barcodeReader.Read())
                Console.WriteLine("Codetext found: " + barcodeReader.GetCodeText() + ", Symbology: " + barcodeReader.GetCodeType());

            // Close the reader
            barcodeReader.Close();

            return barcodeReader;
        }

        [Test]
        [Ignore("WORDSNET-13854")]
        public void FieldDatabase()
        {
            //ExStart
            //ExFor:FieldDatabase
            //ExFor:FieldDatabase.Connection
            //ExFor:FieldDatabase.FileName
            //ExFor:FieldDatabase.FirstRecord
            //ExFor:FieldDatabase.FormatAttributes
            //ExFor:FieldDatabase.InsertHeadings
            //ExFor:FieldDatabase.InsertOnceOnMailMerge
            //ExFor:FieldDatabase.LastRecord
            //ExFor:FieldDatabase.Query
            //ExFor:FieldDatabase.TableFormat
            //ExSummary:Shows how to extract data from a database, and insert it as a field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // This DATABASE field will run a query on a database, and display the result in the form of a table.
            FieldDatabase field = (FieldDatabase)builder.InsertField(FieldType.FieldDatabase, true);
            field.FileName = MyDir + @"Database\Northwind.mdb";
            field.Connection = "DSN=MS Access Databases";
            field.Query = "SELECT * FROM [Products]";

            Assert.AreEqual($" DATABASE  \\d \"{DatabaseDir.Replace("\\", "\\\\") + "Northwind.mdb"}\" \\c \"DSN=MS Access Databases\" \\s \"SELECT * FROM [Products]\"", 
                field.GetFieldCode());

            // Insert another DATABASE field with a more complex query which sorts all products in descending order by gross sales.
            field = (FieldDatabase)builder.InsertField(FieldType.FieldDatabase, true);
            field.FileName = MyDir + @"Database\Northwind.mdb";
            field.Connection = "DSN=MS Access Databases";
            field.Query =
                "SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
                "FROM([Products] " +
                "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
                "GROUP BY[Products].ProductName " +
                "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC";

            // These attributes have the same function as LIMIT and TOP clauses.
            // Configure them to display only rows 1 to 10 of the query result in the field's table.
            field.FirstRecord = "1";
            field.LastRecord = "10";

            // This attribute is the index of the format we want to use for our table. The list of table formats is in the "Table AutoFormat..." menu
            // that shows up when we create a DATABASE field in Microsoft Word. Index #10 corresponds to the "Colorful 3" format.
            field.TableFormat = "10";

            // This attribute decides which elements of the table format we picked above are incorporated into our table.
            // The number we use is the sum of a combination of values corresponding to different aspects of the table style.
            // 63 represents 1 (borders) + 2 (shading) + 4 (font) + 8 (color) + 16 (autofit) + 32 (heading rows).
            field.FormatAttributes = "63";
            field.InsertHeadings = true;
            field.InsertOnceOnMailMerge = true;

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.DATABASE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.DATABASE.docx");

            Assert.AreEqual(2, doc.Range.Fields.Count);
            
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(77, table.Rows.Count);
            Assert.AreEqual(10, table.Rows[0].Cells.Count);

            field = (FieldDatabase)doc.Range.Fields[0];

            Assert.AreEqual($" DATABASE  \\d \"{DatabaseDir.Replace("\\", "\\\\") + "Northwind.mdb"}\" \\c \"DSN=MS Access Databases\" \\s \"SELECT * FROM [Products]\"",
                field.GetFieldCode());

            TestUtil.TableMatchesQueryResult(table, DatabaseDir + "Northwind.mdb", field.Query);

            table = (Table)doc.GetChild(NodeType.Table, 1, true);
            field = (FieldDatabase)doc.Range.Fields[1];

            Assert.AreEqual(11, table.Rows.Count);
            Assert.AreEqual(2, table.Rows[0].Cells.Count);
            Assert.AreEqual("ProductName\a", table.Rows[0].Cells[0].GetText());
            Assert.AreEqual("GrossSales\a", table.Rows[0].Cells[1].GetText());

            Assert.AreEqual($" DATABASE  \\d \"{DatabaseDir.Replace("\\", "\\\\") + "Northwind.mdb"}\" \\c \"DSN=MS Access Databases\" " +
                            $"\\s \"SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
                            "FROM([Products] " +
                            "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
                            "GROUP BY[Products].ProductName " +
                            "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC\" \\f 1 \\t 10 \\l 10 \\b 63 \\h \\o",
                field.GetFieldCode());

            table.Rows[0].Remove();

            TestUtil.TableMatchesQueryResult(table, DatabaseDir + "Northwind.mdb", field.Query.Insert(7, " TOP 10 "));
        }
#endif

        [TestCase(false)]
        [TestCase(true)]
        public void PreserveIncludePicture(bool preserveIncludePictureField)
        {
            //ExStart
            //ExFor:Field.Update(bool)
            //ExFor:LoadOptions.PreserveIncludePictureField
            //ExSummary:Shows how to preserve or discard INCLUDEPICTURE fields when loading a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldIncludePicture includePicture = (FieldIncludePicture)builder.InsertField(FieldType.FieldIncludePicture, true);
            includePicture.SourceFullName = ImageDir + "Transparent background logo.png";
            includePicture.Update(true);

            using (MemoryStream docStream = new MemoryStream())
            {
                doc.Save(docStream, new OoxmlSaveOptions(SaveFormat.Docx));

                // We can set a flag in a LoadOptions object to decide whether to convert all INCLUDEPICTURE fields
                // into image shapes when loading a document that contains them.
                LoadOptions loadOptions = new LoadOptions
                {
                    PreserveIncludePictureField = preserveIncludePictureField
                };

                doc = new Document(docStream, loadOptions);

                if (preserveIncludePictureField)
                {
                    Assert.True(doc.Range.Fields.Any(f => f.Type == FieldType.FieldIncludePicture));

                    doc.UpdateFields();
                    doc.Save(ArtifactsDir + "Field.PreserveIncludePicture.docx");
                }
                else
                {
                    Assert.False(doc.Range.Fields.Any(f => f.Type == FieldType.FieldIncludePicture));
                }
            }
            //ExEnd
        }

        [Test]
        public void FieldFormat()
        {
            //ExStart
            //ExFor:Field.Format
            //ExFor:Field.Update
            //ExFor:FieldFormat
            //ExFor:FieldFormat.DateTimeFormat
            //ExFor:FieldFormat.NumericFormat
            //ExFor:FieldFormat.GeneralFormats
            //ExFor:GeneralFormat
            //ExFor:GeneralFormatCollection
            //ExFor:GeneralFormatCollection.Add(GeneralFormat)
            //ExFor:GeneralFormatCollection.Count
            //ExFor:GeneralFormatCollection.Item(Int32)
            //ExFor:GeneralFormatCollection.Remove(GeneralFormat)
            //ExFor:GeneralFormatCollection.RemoveAt(Int32)
            //ExFor:GeneralFormatCollection.GetEnumerator
            //ExSummary:Shows how to format field results.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a field which displays a result with no format applied.
            Field field = builder.InsertField("= 2 + 3");

            Assert.AreEqual("= 2 + 3", field.GetFieldCode());
            Assert.AreEqual("5", field.Result);

            // We can apply a format to a field's result using the field's attributes.
            // Below are three types of formats that can be applied to a field's result.
            // 1 -  Numeric format:
            FieldFormat format = field.Format;
            format.NumericFormat = "$###.00";
            field.Update();

            Assert.AreEqual("= 2 + 3 \\# $###.00", field.GetFieldCode());
            Assert.AreEqual("$  5.00", field.Result);

            // 2 -  Date/time format:
            field = builder.InsertField("DATE");
            format = field.Format;
            format.DateTimeFormat = "dddd, MMMM dd, yyyy";
            field.Update();

            Assert.AreEqual("DATE \\@ \"dddd, MMMM dd, yyyy\"", field.GetFieldCode());
            Console.WriteLine($"Today's date, in {format.DateTimeFormat} format:\n\t{field.Result}");

            // 3 -  General format:
            field = builder.InsertField("= 25 + 33");
            format = field.Format;
            format.GeneralFormats.Add(GeneralFormat.LowercaseRoman);
            format.GeneralFormats.Add(GeneralFormat.Upper);
            field.Update();

            int index = 0;
            using (IEnumerator<GeneralFormat> generalFormatEnumerator = format.GeneralFormats.GetEnumerator())
                while (generalFormatEnumerator.MoveNext())
                    Console.WriteLine($"General format index {index++}: {generalFormatEnumerator.Current}");

            Assert.AreEqual("= 25 + 33 \\* roman \\* Upper", field.GetFieldCode());
            Assert.AreEqual("LVIII", field.Result);
            Assert.AreEqual(2, format.GeneralFormats.Count);
            Assert.AreEqual(GeneralFormat.LowercaseRoman, format.GeneralFormats[0]);

            // We can remove our formats to revert the field's result to its original form.
            format.GeneralFormats.Remove(GeneralFormat.LowercaseRoman);
            format.GeneralFormats.RemoveAt(0);
            Assert.AreEqual(0, format.GeneralFormats.Count);
            field.Update();

            Assert.AreEqual("= 25 + 33  ", field.GetFieldCode());
            Assert.AreEqual("58", field.Result);
            Assert.AreEqual(0, format.GeneralFormats.Count);
            //ExEnd
        }

        [Test]
        public void Unlink()
        {
            //ExStart
            //ExFor:Document.UnlinkFields
            //ExSummary:Shows how to unlink all fields in the document.
            Document doc = new Document(MyDir + "Linked fields.docx");

            doc.UnlinkFields();
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            string paraWithFields = DocumentHelper.GetParagraphText(doc, 0);

            Assert.AreEqual("Fields.Docx   Элементы указателя не найдены.     1.\r", paraWithFields);
        }

        [Test]
        public void UnlinkAllFieldsInRange()
        {
            //ExStart
            //ExFor:Range.UnlinkFields
            //ExSummary:Shows how to unlink all fields in a range.
            Document doc = new Document(MyDir + "Linked fields.docx");

            Section newSection = (Section)doc.Sections[0].Clone(true);
            doc.Sections.Add(newSection);

            doc.Sections[1].Range.UnlinkFields();
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            string secWithFields = DocumentHelper.GetSectionText(doc, 1);

            Assert.True(secWithFields.Trim().EndsWith(
                "Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4."));
        }

        [Test]
        public void UnlinkSingleField()
        {
            //ExStart
            //ExFor:Field.Unlink
            //ExSummary:Shows how to unlink a field.
            Document doc = new Document(MyDir + "Linked fields.docx");
            doc.Range.Fields[1].Unlink();
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            string paraWithFields = DocumentHelper.GetParagraphText(doc, 0);

            Assert.True(paraWithFields.Trim().EndsWith(
                "FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015"));
        }

        [Test]
        public void UpdateTocPageNumbers()
        {
            Document doc = new Document(MyDir + "Field sample - TOC.docx");

            Node startNode = DocumentHelper.GetParagraph(doc, 2);
            Node endNode = null;

            NodeCollection paragraphCollection = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph para in paragraphCollection.OfType<Paragraph>())
            {
                // Check all runs in the paragraph for the first page breaks
                foreach (Run run in para.Runs.OfType<Run>())
                {
                    if (run.Text.Contains(ControlChar.PageBreak))
                    {
                        endNode = run;
                        break;
                    }
                }
            }

            if (startNode != null && endNode != null)
            {
                RemoveSequence(startNode, endNode);

                startNode.Remove();
                endNode.Remove();
            }

            NodeCollection fStart = doc.GetChildNodes(NodeType.FieldStart, true);

            foreach (FieldStart field in fStart.OfType<FieldStart>())
            {
                FieldType fType = field.FieldType;
                if (fType == FieldType.FieldTOC)
                {
                    Paragraph para = (Paragraph)field.GetAncestor(NodeType.Paragraph);
                    para.Range.UpdateFields();
                    break;
                }
            }

            doc.Save(ArtifactsDir + "Field.UpdateTocPageNumbers.docx");
        }

        private static void RemoveSequence(Node start, Node end)
        {
            Node curNode = start.NextPreOrder(start.Document);
            while (curNode != null && !curNode.Equals(end))
            {
                // Move to next node
                Node nextNode = curNode.NextPreOrder(start.Document);

                // Check whether current contains end node
                if (curNode.IsComposite)
                {
                    CompositeNode curComposite = (CompositeNode)curNode;
                    if (!curComposite.GetChildNodes(NodeType.Any, true).Contains(end) &&
                        !curComposite.GetChildNodes(NodeType.Any, true).Contains(start))
                    {
                        nextNode = curNode.NextSibling;
                        curNode.Remove();
                    }
                }
                else
                {
                    curNode.Remove();
                }

                curNode = nextNode;
            }
        }
        
        //ExStart
        //ExFor:Fields.FieldAsk
        //ExFor:Fields.FieldAsk.BookmarkName
        //ExFor:Fields.FieldAsk.DefaultResponse
        //ExFor:Fields.FieldAsk.PromptOnceOnMailMerge
        //ExFor:Fields.FieldAsk.PromptText
        //ExFor:FieldOptions.UserPromptRespondent
        //ExFor:IFieldUserPromptRespondent
        //ExFor:IFieldUserPromptRespondent.Respond(String,String)
        //ExSummary:Shows how to create an ASK field, and set its properties.
        [Test]
        public void FieldAsk()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Place a field where the response to our ASK field will be placed.
            FieldRef fieldRef = (FieldRef)builder.InsertField(FieldType.FieldRef, true);
            fieldRef.BookmarkName = "MyAskField";
            builder.Writeln();

            Assert.AreEqual(" REF  MyAskField", fieldRef.GetFieldCode());

            // Insert the ASK field and edit its properties, making sure to reference our REF field by bookmark name.
            FieldAsk fieldAsk = (FieldAsk)builder.InsertField(FieldType.FieldAsk, true);
            fieldAsk.BookmarkName = "MyAskField";
            fieldAsk.PromptText = "Please provide a response for this ASK field";
            fieldAsk.DefaultResponse = "Response from within the field.";
            fieldAsk.PromptOnceOnMailMerge = true;
            builder.Writeln();

            Assert.AreEqual(
                " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o",
                fieldAsk.GetFieldCode());

            // ASK fields apply the default response to their respective REF fields during a mail merge.
            DataTable table = new DataTable("My Table");
            table.Columns.Add("Column 1");
            table.Rows.Add("Row 1");
            table.Rows.Add("Row 2");

            FieldMergeField fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Column 1";

            // We can modify or override the default response in our ASK fields with a custom prompt responder,
            // which will take place during a mail merge.
            doc.FieldOptions.UserPromptRespondent = new MyPromptRespondent();
            doc.MailMerge.Execute(table);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.ASK.docx");
            TestFieldAsk(table, doc); //ExSkip
        }

        /// <summary>
        /// Prepends text to the default response of an ASK field during a mail merge.
        /// </summary>
        private class MyPromptRespondent : IFieldUserPromptRespondent
        {
            public string Respond(string promptText, string defaultResponse)
            {
                return "Response from MyPromptRespondent. " + defaultResponse;
            }
        }
        //ExEnd

        private void TestFieldAsk(DataTable dataTable, Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);

            FieldRef fieldRef = (FieldRef)doc.Range.Fields.First(f => f.Type == FieldType.FieldRef);
            TestUtil.VerifyField(FieldType.FieldRef, 
                " REF  MyAskField", "Response from MyPromptRespondent. Response from within the field.", fieldRef);

            FieldAsk fieldAsk = (FieldAsk)doc.Range.Fields.First(f => f.Type == FieldType.FieldAsk);
            TestUtil.VerifyField(FieldType.FieldAsk, 
                " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o", 
                "Response from MyPromptRespondent. Response from within the field.", fieldAsk);
            
            Assert.AreEqual("MyAskField", fieldAsk.BookmarkName);
            Assert.AreEqual("Please provide a response for this ASK field", fieldAsk.PromptText);
            Assert.AreEqual("Response from within the field.", fieldAsk.DefaultResponse);
            Assert.AreEqual(true, fieldAsk.PromptOnceOnMailMerge);

            TestUtil.MailMergeMatchesDataTable(dataTable, doc, true);
        }

        [Test]
        public void FieldAdvance()
        {
            //ExStart
            //ExFor:Fields.FieldAdvance
            //ExFor:Fields.FieldAdvance.DownOffset
            //ExFor:Fields.FieldAdvance.HorizontalPosition
            //ExFor:Fields.FieldAdvance.LeftOffset
            //ExFor:Fields.FieldAdvance.RightOffset
            //ExFor:Fields.FieldAdvance.UpOffset
            //ExFor:Fields.FieldAdvance.VerticalPosition
            //ExSummary:Shows how to insert an ADVANCE field, and edit its properties. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("This text is in its normal place.");

            // Below are two ways of using the ADVANCE field to adjust the position of text that follows it.
            // The effects of an ADVANCE field continue to be applied until the paragraph ends,
            // or another ADVANCE field updates the offset/coordinate values.
            // 1 -  Specify a directional offset:
            FieldAdvance field = (FieldAdvance)builder.InsertField(FieldType.FieldAdvance, true);
            Assert.AreEqual(FieldType.FieldAdvance, field.Type); //ExSkip
            Assert.AreEqual(" ADVANCE ", field.GetFieldCode()); //ExSkip
            field.RightOffset = "5";
            field.UpOffset = "5";

            Assert.AreEqual(" ADVANCE  \\r 5 \\u 5", field.GetFieldCode());

            builder.Write("This text will be moved up and to the right.");
            
            field = (FieldAdvance)builder.InsertField(FieldType.FieldAdvance, true);
            field.DownOffset = "5";
            field.LeftOffset = "100";

            Assert.AreEqual(" ADVANCE  \\d 5 \\l 100", field.GetFieldCode());

            builder.Writeln("This text is moved down and to the left, overlapping the previous text.");

            // 2 -  Move text to a position specified by coordinates:
            field = (FieldAdvance)builder.InsertField(FieldType.FieldAdvance, true);
            field.HorizontalPosition = "-100";
            field.VerticalPosition = "200";

            Assert.AreEqual(" ADVANCE  \\x -100 \\y 200", field.GetFieldCode());

            builder.Write("This text is in a custom position.");

            doc.Save(ArtifactsDir + "Field.ADVANCE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.ADVANCE.docx");

            field = (FieldAdvance)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldAdvance, " ADVANCE  \\r 5 \\u 5", string.Empty, field);
            Assert.AreEqual("5", field.RightOffset);
            Assert.AreEqual("5", field.UpOffset);

            field = (FieldAdvance)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldAdvance, " ADVANCE  \\d 5 \\l 100", string.Empty, field);
            Assert.AreEqual("5", field.DownOffset);
            Assert.AreEqual("100", field.LeftOffset);

            field = (FieldAdvance)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldAdvance, " ADVANCE  \\x -100 \\y 200", string.Empty, field);
            Assert.AreEqual("-100", field.HorizontalPosition);
            Assert.AreEqual("200", field.VerticalPosition);
        }

        [Test]
        public void FieldAddressBlock()
        {
            //ExStart
            //ExFor:Fields.FieldAddressBlock.ExcludedCountryOrRegionName
            //ExFor:Fields.FieldAddressBlock.FormatAddressOnCountryOrRegion
            //ExFor:Fields.FieldAddressBlock.IncludeCountryOrRegionName
            //ExFor:Fields.FieldAddressBlock.LanguageId
            //ExFor:Fields.FieldAddressBlock.NameAndAddressFormat
            //ExSummary:Shows how to insert an ADDRESSBLOCK field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldAddressBlock field = (FieldAddressBlock)builder.InsertField(FieldType.FieldAddressBlock, true);

            Assert.AreEqual(" ADDRESSBLOCK ", field.GetFieldCode());

            // Setting this to "2" will cause all countries/regions to be included,
            // unless it is the one specified in the ExcludedCountryOrRegionName attribute.
            field.IncludeCountryOrRegionName = "2";
            field.FormatAddressOnCountryOrRegion = true;
            field.ExcludedCountryOrRegionName = "United States";
            field.NameAndAddressFormat = "<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>";

            // By default, the language ID will be set to that of the first character of the document.
            // We can set a culture for the field to format the result with like this.
            field.LanguageId = new CultureInfo("en-US").LCID.ToString();

            Assert.AreEqual(
                " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
                field.GetFieldCode());
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            field = (FieldAddressBlock)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldAddressBlock, 
                " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033", 
                "«AddressBlock»", field);
            Assert.AreEqual("2", field.IncludeCountryOrRegionName);
            Assert.AreEqual(true, field.FormatAddressOnCountryOrRegion);
            Assert.AreEqual("United States", field.ExcludedCountryOrRegionName);
            Assert.AreEqual("<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>",
                field.NameAndAddressFormat);
            Assert.AreEqual("1033", field.LanguageId);
        }

        //ExStart
        //ExFor:FieldCollection
        //ExFor:FieldCollection.Count
        //ExFor:FieldCollection.GetEnumerator
        //ExFor:FieldStart
        //ExFor:FieldStart.Accept(DocumentVisitor)
        //ExFor:FieldSeparator
        //ExFor:FieldSeparator.Accept(DocumentVisitor)
        //ExFor:FieldEnd
        //ExFor:FieldEnd.Accept(DocumentVisitor)
        //ExFor:FieldEnd.HasSeparator
        //ExFor:Field.End
        //ExFor:Field.Separator
        //ExFor:Field.Start
        //ExSummary:Shows how to work with a collection of fields.
        [Test] //ExSkip
        public void FieldCollection()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(" DATE \\@ \"dddd, d MMMM yyyy\" ");
            builder.InsertField(" TIME ");
            builder.InsertField(" REVNUM ");
            builder.InsertField(" AUTHOR  \"John Doe\" ");
            builder.InsertField(" SUBJECT \"My Subject\" ");
            builder.InsertField(" QUOTE \"Hello world!\" ");
            doc.UpdateFields();

            // This collection stores all of a document's fields.
            FieldCollection fields = doc.Range.Fields;

            Assert.AreEqual(6, fields.Count);

            // Iterate over the field collection, and print contents and type
            // of every field using a custom visitor implementation.
            FieldVisitor fieldVisitor = new FieldVisitor();

            using (IEnumerator<Field> fieldEnumerator = fields.GetEnumerator())
            {
                while (fieldEnumerator.MoveNext())
                {
                    if (fieldEnumerator.Current != null)
                    {
                        fieldEnumerator.Current.Start.Accept(fieldVisitor);
                        fieldEnumerator.Current.Separator?.Accept(fieldVisitor);
                        fieldEnumerator.Current.End.Accept(fieldVisitor);
                    }
                    else
                    {
                        Console.WriteLine("There are no fields in the document.");
                    }
                }
            }

            Console.WriteLine(fieldVisitor.GetText());
            TestFieldCollection(fieldVisitor.GetText()); //ExSkip
        }

        /// <summary>
        /// Document visitor implementation that prints field info.
        /// </summary>
        public class FieldVisitor : DocumentVisitor
        {
            public FieldVisitor()
            {
                mBuilder = new StringBuilder();
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a FieldStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldStart(FieldStart fieldStart)
            {
                mBuilder.AppendLine("Found field: " + fieldStart.FieldType);
                mBuilder.AppendLine("\tField code: " + fieldStart.GetField().GetFieldCode());
                mBuilder.AppendLine("\tDisplayed as: " + fieldStart.GetField().Result);

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                mBuilder.AppendLine("\tFound separator: " + fieldSeparator.GetText());

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                mBuilder.AppendLine("End of field: " + fieldEnd.FieldType);

                return VisitorAction.Continue;
            }

            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private void TestFieldCollection(string fieldVisitorText)
        {
            Assert.True(fieldVisitorText.Contains("Found field: FieldDate"));
            Assert.True(fieldVisitorText.Contains("Found field: FieldTime"));
            Assert.True(fieldVisitorText.Contains("Found field: FieldRevisionNum"));
            Assert.True(fieldVisitorText.Contains("Found field: FieldAuthor"));
            Assert.True(fieldVisitorText.Contains("Found field: FieldSubject"));
            Assert.True(fieldVisitorText.Contains("Found field: FieldQuote"));
        }

        [Test]
        public void RemoveFields()
        {
            //ExStart
            //ExFor:FieldCollection
            //ExFor:FieldCollection.Count
            //ExFor:FieldCollection.Clear
            //ExFor:FieldCollection.Item(Int32)
            //ExFor:FieldCollection.Remove(Field)
            //ExFor:FieldCollection.Remove(FieldStart)
            //ExFor:FieldCollection.RemoveAt(Int32)
            //ExFor:Field.Remove
            //ExSummary:Shows how to remove fields from a field collection.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(" DATE \\@ \"dddd, d MMMM yyyy\" ");
            builder.InsertField(" TIME ");
            builder.InsertField(" REVNUM ");
            builder.InsertField(" AUTHOR  \"John Doe\" ");
            builder.InsertField(" SUBJECT \"My Subject\" ");
            builder.InsertField(" QUOTE \"Hello world!\" ");
            doc.UpdateFields();

            // This collection stores all of a document's fields.
            FieldCollection fields = doc.Range.Fields;

            Assert.AreEqual(6, fields.Count);

            // Below are four ways of removing fields from a field collection.
            // 1 -  Get a field to remove itself.
            fields[0].Remove();
            Assert.AreEqual(5, fields.Count);

            // 2 -  Get the collection to remove a field that we pass to its removal method.
            Field lastField = fields[3];
            fields.Remove(lastField);
            Assert.AreEqual(4, fields.Count);

            // 3 -  Remove a field from a collection at an index.
            fields.RemoveAt(2);
            Assert.AreEqual(3, fields.Count);

            // 4 -  Remove all the fields from the collection at once.
            fields.Clear();
            Assert.AreEqual(0, fields.Count);
            //ExEnd
        }

        [Test]
        public void FieldCompare()
        {
            //ExStart
            //ExFor:FieldCompare
            //ExFor:FieldCompare.ComparisonOperator
            //ExFor:FieldCompare.LeftExpression
            //ExFor:FieldCompare.RightExpression
            //ExSummary:Shows how to compare expressions using a COMPARE field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldCompare field = (FieldCompare)builder.InsertField(FieldType.FieldCompare, true);
            field.LeftExpression = "3";
            field.ComparisonOperator = "<";
            field.RightExpression = "2";
            field.Update();

            // The COMPARE field displays a "0", or a "1", depending on the truth of its statement.
            // The result of this statement is false, so this field will display a "0".
            Assert.AreEqual(" COMPARE  3 < 2", field.GetFieldCode());
            Assert.AreEqual("0", field.Result);

            builder.Writeln();

            field = (FieldCompare)builder.InsertField(FieldType.FieldCompare, true);
            field.LeftExpression = "5";
            field.ComparisonOperator = "=";
            field.RightExpression = "2 + 3";
            field.Update();

            // This field displays a "1" since the statement is true.
            Assert.AreEqual(" COMPARE  5 = \"2 + 3\"", field.GetFieldCode());
            Assert.AreEqual("1", field.Result);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.COMPARE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.COMPARE.docx");

            field = (FieldCompare)doc.Range.Fields[0];
            
            TestUtil.VerifyField(FieldType.FieldCompare, " COMPARE  3 < 2", "0", field);
            Assert.AreEqual("3", field.LeftExpression);
            Assert.AreEqual("<", field.ComparisonOperator);
            Assert.AreEqual("2", field.RightExpression);

            field = (FieldCompare)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldCompare, " COMPARE  5 = \"2 + 3\"", "1", field);
            Assert.AreEqual("5", field.LeftExpression);
            Assert.AreEqual("=", field.ComparisonOperator);
            Assert.AreEqual("\"2 + 3\"", field.RightExpression);
        }

        [Test]
        public void FieldIf()
        {
            //ExStart
            //ExFor:FieldIf
            //ExFor:FieldIf.ComparisonOperator
            //ExFor:FieldIf.EvaluateCondition
            //ExFor:FieldIf.FalseText
            //ExFor:FieldIf.LeftExpression
            //ExFor:FieldIf.RightExpression
            //ExFor:FieldIf.TrueText
            //ExFor:FieldIfComparisonResult
            //ExSummary:Shows how to insert an IF field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Statement 1: ");
            FieldIf field = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
            field.LeftExpression = "0";
            field.ComparisonOperator = "=";
            field.RightExpression = "1";

            // The IF field will display a string from either its "TrueText" attribute,
            // or its "FalseText" attribute, depending on the truth of the statement that we have constructed.
            field.TrueText = "True";
            field.FalseText = "False";
            field.Update();

            // In this case, "0 = 1" is incorrect, so the displayed result will be "False".
            Assert.AreEqual(" IF  0 = 1 True False", field.GetFieldCode());
            Assert.AreEqual(FieldIfComparisonResult.False, field.EvaluateCondition());
            Assert.AreEqual("False", field.Result);

            builder.Write("\nStatement 2: ");
            field = (FieldIf)builder.InsertField(FieldType.FieldIf, true);
            field.LeftExpression = "5";
            field.ComparisonOperator = "=";
            field.RightExpression = "2 + 3";
            field.TrueText = "True";
            field.FalseText = "False";
            field.Update();

            // This time the statement is correct, so the displayed result will be "True".
            Assert.AreEqual(" IF  5 = \"2 + 3\" True False", field.GetFieldCode());
            Assert.AreEqual(FieldIfComparisonResult.True, field.EvaluateCondition());
            Assert.AreEqual("True", field.Result);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.IF.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.IF.docx");
            field = (FieldIf)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldIf, " IF  0 = 1 True False", "False", field);
            Assert.AreEqual("0", field.LeftExpression);
            Assert.AreEqual("=", field.ComparisonOperator);
            Assert.AreEqual("1", field.RightExpression);
            Assert.AreEqual("True", field.TrueText);
            Assert.AreEqual("False", field.FalseText);

            field = (FieldIf)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldIf, " IF  5 = \"2 + 3\" True False", "True", field);
            Assert.AreEqual("5", field.LeftExpression);
            Assert.AreEqual("=", field.ComparisonOperator);
            Assert.AreEqual("\"2 + 3\"", field.RightExpression);
            Assert.AreEqual("True", field.TrueText);
            Assert.AreEqual("False", field.FalseText);
        }

        [Test]
        public void FieldAutoNum()
        {
            //ExStart
            //ExFor:FieldAutoNum
            //ExFor:FieldAutoNum.SeparatorCharacter
            //ExSummary:Shows how to number paragraphs using autonum fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // AUTONUM fields display a number that increments at each AUTONUM field,
            // which allows us to automatically number items similar to a numbered list.
            // This field will display a number "1.".
            FieldAutoNum field = (FieldAutoNum)builder.InsertField(FieldType.FieldAutoNum, true);
            builder.Writeln("\tParagraph 1.");

            Assert.AreEqual(" AUTONUM ", field.GetFieldCode());

            field = (FieldAutoNum)builder.InsertField(FieldType.FieldAutoNum, true);
            builder.Writeln("\tParagraph 2.");

            // The separator character, which appears in the field result immediately after the number,
            // is a full stop by default. If we leave this attribute null, our second AUTONUM field will display "2." in the document.
            Assert.IsNull(field.SeparatorCharacter);

            // We can set this attribute to apply the first character of its string as the new separator character.
            // In this case, our AUTONUM field will now display "2:".
            field.SeparatorCharacter = ":";

            Assert.AreEqual(" AUTONUM  \\s :", field.GetFieldCode());

            doc.Save(ArtifactsDir + "Field.AUTONUM.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.AUTONUM.docx");

            TestUtil.VerifyField(FieldType.FieldAutoNum, " AUTONUM ", string.Empty, doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldAutoNum, " AUTONUM  \\s :", string.Empty, doc.Range.Fields[1]);
        }

        //ExStart
        //ExFor:FieldAutoNumLgl
        //ExFor:FieldAutoNumLgl.RemoveTrailingPeriod
        //ExFor:FieldAutoNumLgl.SeparatorCharacter
        //ExSummary:Shows how to organize a document using AUTONUMLGL fields.
        [Test] //ExSkip
        public void FieldAutoNumLgl()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            const string fillerText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                                      "\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

            // AUTONUMLGL fields display a number that increments at each AUTONUMLGL field within its current heading level.
            // These fields maintain a separate count for each heading level,
            // and each field also displays the AUTONUMLGL field counts for all heading levels below its own. 
            // If the count for any heading level changes, the counts for all levels above that level are reset to 1.
            // This allows us to organize our document in the form of an outline list.
            // This is the first AUTONUMLGL field at a heading level of 1, so it will display "1." in the document.
            InsertNumberedClause(builder, "\tHeading 1", fillerText, StyleIdentifier.Heading1);

            // This is the second AUTONUMLGL field at a heading level of 1, so it will display "2.".
            InsertNumberedClause(builder, "\tHeading 2", fillerText, StyleIdentifier.Heading1);

            // This is the first AUTONUMLGL field at a heading level of 2,
            // and the AUTONUMLGL count for the heading level below it is "2", so it will display "2.1.".
            InsertNumberedClause(builder, "\tHeading 3", fillerText, StyleIdentifier.Heading2);

            // This is the first AUTONUMLGL field at a heading level of 3. 
            // Working in the same way as the field above, it will display "2.1.1.".
            InsertNumberedClause(builder, "\tHeading 4", fillerText, StyleIdentifier.Heading3);

            // This field is at a heading level of 2, and its respective AUTONUMLGL count is at 2, so the field will display "2.2.".
            InsertNumberedClause(builder, "\tHeading 5", fillerText, StyleIdentifier.Heading2);

            // Incrementing the AUTONUMLGL count for a heading level below this one
            // has reset the count for this level, so this field will display "2.2.1.".
            InsertNumberedClause(builder, "\tHeading 6", fillerText, StyleIdentifier.Heading3);

            foreach (FieldAutoNumLgl field in doc.Range.Fields.Where(f => f.Type == FieldType.FieldAutoNumLegal))
            {
                // The separator character, which appears in the field result immediately after the number,
                // is a full stop by default. If we leave this attribute null,
                // our last AUTONUMLGL field will display "2.2.1." in the document.
                Assert.IsNull(field.SeparatorCharacter);

                // Setting a custom separater character and removing the trailing period
                // will change that field's appearance from "2.2.1." to "2:2:1".
                // We will apply this to all the fields that we have created.
                field.SeparatorCharacter = ":";
                field.RemoveTrailingPeriod = true;
                Assert.AreEqual(" AUTONUMLGL  \\s : \\e", field.GetFieldCode());
            }

            doc.Save(ArtifactsDir + "Field.AUTONUMLGL.docx");
            TestFieldAutoNumLgl(doc); //ExSkip
        }

        /// <summary>
        /// Uses a document builder to insert a clause numbered by an AUTONUMLGL field.
        /// </summary>
        private static void InsertNumberedClause(DocumentBuilder builder, string heading, string contents, StyleIdentifier headingStyle)
        {
            builder.InsertField(FieldType.FieldAutoNumLegal, true);
            builder.CurrentParagraph.ParagraphFormat.StyleIdentifier = headingStyle;
            builder.Writeln(heading);

            // This text will belong to the auto num legal field above it.
            // It will collapse when we click the arrow next to the corresponding AUTONUMLGL field in Microsoft Word.
            builder.CurrentParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.BodyText;
            builder.Writeln(contents);
        }
        //ExEnd

        private void TestFieldAutoNumLgl(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);

            foreach (FieldAutoNumLgl field in doc.Range.Fields.Where(f => f.Type == FieldType.FieldAutoNumLegal))
            {
                TestUtil.VerifyField(FieldType.FieldAutoNumLegal, " AUTONUMLGL  \\s : \\e", string.Empty, field);
                
                Assert.AreEqual(":", field.SeparatorCharacter);
                Assert.True(field.RemoveTrailingPeriod);
            }
        }

        [Test]
        public void FieldAutoNumOut()
        {
            //ExStart
            //ExFor:FieldAutoNumOut
            //ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // AUTONUMOUT fields display a number that increments at each AUTONUMOUT field.
            // Unlike AUTONUM fields, AUTONUMOUT fields use the outline numbering scheme,
            // which we can define in Microsoft Word via Format -> Bullets & Numbering -> "Outline Numbered".
            // This allows us to automatically number items similar to a numbered list.
            // LISTNUM fields are a newer alternative to AUTONUMOUT fields.
            // This field will display "1.".
            builder.InsertField(FieldType.FieldAutoNumOutline, true);
            builder.Writeln("\tParagraph 1.");

            // This field will display "2.".
            builder.InsertField(FieldType.FieldAutoNumOutline, true);
            builder.Writeln("\tParagraph 2.");

            foreach (FieldAutoNumOut field in doc.Range.Fields.Where(f => f.Type == FieldType.FieldAutoNumOutline))
                Assert.AreEqual(" AUTONUMOUT ", field.GetFieldCode());

            doc.Save(ArtifactsDir + "Field.AUTONUMOUT.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.AUTONUMOUT.docx");

            foreach (Field field in doc.Range.Fields)
                TestUtil.VerifyField(FieldType.FieldAutoNumOutline, " AUTONUMOUT ", string.Empty, field);
        }

        [Test]
        public void FieldAutoText()
        {
            //ExStart
            //ExFor:Fields.FieldAutoText
            //ExFor:FieldAutoText.EntryName
            //ExFor:FieldOptions.BuiltInTemplatesPaths
            //ExFor:FieldGlossary
            //ExFor:FieldGlossary.EntryName
            //ExSummary:Shows how to display a building block with AUTOTEXT and GLOSSARY fields. 
            Document doc = new Document();

            // Create a glossary document, and add an AutoText building block to it.
            doc.GlossaryDocument = new GlossaryDocument();
            BuildingBlock buildingBlock = new BuildingBlock(doc.GlossaryDocument);
            buildingBlock.Name = "MyBlock";
            buildingBlock.Gallery = BuildingBlockGallery.AutoText;
            buildingBlock.Category = "General";
            buildingBlock.Description = "MyBlock description";
            buildingBlock.Behavior = BuildingBlockBehavior.Paragraph;
            doc.GlossaryDocument.AppendChild(buildingBlock);

            // Create a source, and add it as text to our building block.
            Document buildingBlockSource = new Document();
            DocumentBuilder buildingBlockSourceBuilder = new DocumentBuilder(buildingBlockSource);
            buildingBlockSourceBuilder.Writeln("Hello World!");

            Node buildingBlockContent = doc.GlossaryDocument.ImportNode(buildingBlockSource.FirstSection, true);
            buildingBlock.AppendChild(buildingBlockContent);

            // Set a file which contains parts that our document, or its attached template may not contain.
            doc.FieldOptions.BuiltInTemplatesPaths = new[] { MyDir + "Busniess brochure.dotx" };

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two ways to use fields to display the contents of our building block.
            // 1 -  Using an AUTOTEXT field:
            FieldAutoText fieldAutoText = (FieldAutoText)builder.InsertField(FieldType.FieldAutoText, true);
            fieldAutoText.EntryName = "MyBlock";

            Assert.AreEqual(" AUTOTEXT  MyBlock", fieldAutoText.GetFieldCode());
            
            // 2 -  Using a GLOSSARY field:
            FieldGlossary fieldGlossary = (FieldGlossary)builder.InsertField(FieldType.FieldGlossary, true);
            fieldGlossary.EntryName = "MyBlock";

            Assert.AreEqual(" GLOSSARY  MyBlock", fieldGlossary.GetFieldCode());

			doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.AUTOTEXT.GLOSSARY.dotx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.AUTOTEXT.GLOSSARY.dotx");
            
            Assert.That(doc.FieldOptions.BuiltInTemplatesPaths, Is.Empty);

            fieldAutoText = (FieldAutoText)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldAutoText, " AUTOTEXT  MyBlock", "Hello World!\r", fieldAutoText);
            Assert.AreEqual("MyBlock", fieldAutoText.EntryName);

            fieldGlossary = (FieldGlossary)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldGlossary, " GLOSSARY  MyBlock", "Hello World!\r", fieldGlossary);
            Assert.AreEqual("MyBlock", fieldGlossary.EntryName);
        }

        //ExStart
        //ExFor:Fields.FieldAutoTextList
        //ExFor:Fields.FieldAutoTextList.EntryName
        //ExFor:Fields.FieldAutoTextList.ListStyle
        //ExFor:Fields.FieldAutoTextList.ScreenTip
        //ExSummary:Shows how to use an AUTOTEXTLIST field to select from a list of AutoText entries.
        [Test] //ExSkip
        public void FieldAutoTextList()
        {
            Document doc = new Document();

            // Create a glossary document, and populate it with auto text entries.
            doc.GlossaryDocument = new GlossaryDocument();
            AppendAutoTextEntry(doc.GlossaryDocument, "AutoText 1", "Contents of AutoText 1");
            AppendAutoTextEntry(doc.GlossaryDocument, "AutoText 2", "Contents of AutoText 2");
            AppendAutoTextEntry(doc.GlossaryDocument, "AutoText 3", "Contents of AutoText 3");

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an AUTOTEXTLIST field, and set the text that the field will display in Microsoft Word.
            // Set the text to prompt the user to right click this field to select an AutoText building block,
            // whose contents the field will then display.
            FieldAutoTextList field = (FieldAutoTextList)builder.InsertField(FieldType.FieldAutoTextList, true);
            field.EntryName = "Right click here to select an AutoText block";
            field.ListStyle = "Heading 1";
            field.ScreenTip = "Hover tip text for AutoTextList goes here";

            Assert.AreEqual(" AUTOTEXTLIST  \"Right click here to select an AutoText block\" " +
                            "\\s \"Heading 1\" " +
                            "\\t \"Hover tip text for AutoTextList goes here\"", field.GetFieldCode());

            doc.Save(ArtifactsDir + "Field.AUTOTEXTLIST.dotx");
            TestFieldAutoTextList(doc); //ExSkip
        }

        /// <summary>
        /// Create an AutoText-type building block, and add it to a glossary document.
        /// </summary>
        private static void AppendAutoTextEntry(GlossaryDocument glossaryDoc, string name, string contents)
        {
            BuildingBlock buildingBlock = new BuildingBlock(glossaryDoc);
            buildingBlock.Name = name;
            buildingBlock.Gallery = BuildingBlockGallery.AutoText;
            buildingBlock.Category = "General";
            buildingBlock.Behavior = BuildingBlockBehavior.Paragraph;

            Section section = new Section(glossaryDoc);
            section.AppendChild(new Body(glossaryDoc));
            section.Body.AppendParagraph(contents);
            buildingBlock.AppendChild(section);

            glossaryDoc.AppendChild(buildingBlock);
        }
        //ExEnd

        private void TestFieldAutoTextList(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual(3, doc.GlossaryDocument.Count);
            Assert.AreEqual("AutoText 1", doc.GlossaryDocument.BuildingBlocks[0].Name);
            Assert.AreEqual("Contents of AutoText 1", doc.GlossaryDocument.BuildingBlocks[0].GetText().Trim());
            Assert.AreEqual("AutoText 2", doc.GlossaryDocument.BuildingBlocks[1].Name);
            Assert.AreEqual("Contents of AutoText 2", doc.GlossaryDocument.BuildingBlocks[1].GetText().Trim());
            Assert.AreEqual("AutoText 3", doc.GlossaryDocument.BuildingBlocks[2].Name);
            Assert.AreEqual("Contents of AutoText 3", doc.GlossaryDocument.BuildingBlocks[2].GetText().Trim());

            FieldAutoTextList field = (FieldAutoTextList)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldAutoTextList,
                " AUTOTEXTLIST  \"Right click here to select an AutoText block\" \\s \"Heading 1\" \\t \"Hover tip text for AutoTextList goes here\"",
                string.Empty, field);
            Assert.AreEqual("Right click here to select an AutoText block", field.EntryName);
            Assert.AreEqual("Heading 1", field.ListStyle);
            Assert.AreEqual("Hover tip text for AutoTextList goes here", field.ScreenTip);
        }

        [Test]
        public void FieldGreetingLine()
        {
            //ExStart
            //ExFor:FieldGreetingLine
            //ExFor:FieldGreetingLine.AlternateText
            //ExFor:FieldGreetingLine.GetFieldNames
            //ExFor:FieldGreetingLine.LanguageId
            //ExFor:FieldGreetingLine.NameFormat
            //ExSummary:Shows how to insert a GREETINGLINE field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a generic greeting using a GREETINGLINE field, and some text after it.
            FieldGreetingLine field = (FieldGreetingLine)builder.InsertField(FieldType.FieldGreetingLine, true);
            builder.Writeln("\n\n\tThis is your custom greeting, created programmatically using Aspose Words!");

            // A GREETINGLINE field accepts values from a data source during a mail merge, like a MERGEFIELD.
            // It can also format the way in which the data from the source
            // is written in its place once the mail merge is complete.
            // The field names collection corresponds to the columns from the data source
            // that the field will take values from.
            Assert.AreEqual(0, field.GetFieldNames().Length);

            // To populate that array, we need to specify a format for our greeting line.
            field.NameFormat = "<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> ";

            // Now, our field will accept values from these two columns in the data source.
            Assert.AreEqual("Courtesy Title", field.GetFieldNames()[0]);
            Assert.AreEqual("Last Name", field.GetFieldNames()[1]);
            Assert.AreEqual(2, field.GetFieldNames().Length);

            // This string will cover any cases where the data in the data table is invalid
            // by substituting the malformed name with a string.
            field.AlternateText = "Sir or Madam";

            // Set a locale to format the result with.
            field.LanguageId = new CultureInfo("en-US").LCID.ToString();

            Assert.AreEqual(" GREETINGLINE  \\f \"<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> \" \\e \"Sir or Madam\" \\l 1033", 
                field.GetFieldCode());

            // Create a data table with columns whose names match elements
            // from the field's field names collection, and then carry out the mail merge.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Courtesy Title");
            table.Columns.Add("First Name");
            table.Columns.Add("Last Name");
            table.Rows.Add("Mr.", "John", "Doe");
            table.Rows.Add("Mrs.", "Jane", "Cardholder");

            // This row has an invalid value in the Courtesy Title column, so our greeting will default to the alternate text.
            table.Rows.Add("", "No", "Name");

            doc.MailMerge.Execute(table);

            Assert.That(doc.Range.Fields, Is.Empty);
            Assert.AreEqual("Dear Mr. Doe,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
                            "\fDear Mrs. Cardholder,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
                            "\fDear Sir or Madam,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!",
                doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void FieldListNum()
        {
            //ExStart
            //ExFor:FieldListNum
            //ExFor:FieldListNum.HasListName
            //ExFor:FieldListNum.ListLevel
            //ExFor:FieldListNum.ListName
            //ExFor:FieldListNum.StartingNumber
            //ExSummary:Shows how to number paragraphs with LISTNUM fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // LISTNUM fields display a number that increments at each LISTNUM field.
            // These fields also have a variety of options that allow us to use them to emulate numbered lists.
            FieldListNum field = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);

            // Lists start counting at 1 by default, but we can set this number to a different value, such as 0.
            // This field will display "0)".
            field.StartingNumber = "0";
            builder.Writeln("Paragraph 1");

            Assert.AreEqual(" LISTNUM  \\s 0", field.GetFieldCode());

            // LISTNUM fields maintain separate counts for each list level. 
            // Inserting a LISTNUM field in the same paragraph as another LISTNUM field
            // increases the list level instead of the count.
            // The next field will continue the count we started above, and will have a value of 1 at list level 1.
            builder.InsertField(FieldType.FieldListNum, true);

            // This field will start a count at list level 2, and will display a value of 1.
            builder.InsertField(FieldType.FieldListNum, true);

            // This field will start a count at list level 3, and will display a value of 1.
            // Different list levels have different formatting,
            // so these fields combined will display a value of "1)a)i)".
            builder.InsertField(FieldType.FieldListNum, true);
            builder.Writeln("Paragraph 2");

            // The next LISTNUM field that we insert will continue the count at the list level
            // that the previous LISTNUM field was on.
            // We can use the "ListLevel" attribute to jump to a different list level.
            // If this LISTNUM field stayed on list level 3, it would display "ii)",
            // but, since we have moved it to list level 2, it carries on the count at that level and displays "b)".
            field = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);
            field.ListLevel = "2";
            builder.Writeln("Paragraph 3");

            Assert.AreEqual(" LISTNUM  \\l 2", field.GetFieldCode());

            // We can set the ListName attribute to get the field to emulate a different AUTONUM field type.
            // "NumberDefault" emulates AUTONUM, "OutlineDefault" emulates AUTONUMOUT, and "LegalDefault" emulates AUTONUMLGL fields.
            // The "OutlineDefault" list name with 1 as the starting number will result in the field displaying "I.".
            field = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);
            field.StartingNumber = "1";
            field.ListName = "OutlineDefault";
            builder.Writeln("Paragraph 4");

            Assert.IsTrue(field.HasListName);
            Assert.AreEqual(" LISTNUM  OutlineDefault \\s 1", field.GetFieldCode());

            // The ListName does not carry over from the previous field, and needs to be set each time.
            // This field continues the count with the different list name, and displays "II.".
            field = (FieldListNum)builder.InsertField(FieldType.FieldListNum, true);
            field.ListName = "OutlineDefault";
            builder.Writeln("Paragraph 5");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.LISTNUM.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.LISTNUM.docx");

            Assert.AreEqual(7, doc.Range.Fields.Count);

            field = (FieldListNum)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldListNum, " LISTNUM  \\s 0", string.Empty, field);
            Assert.AreEqual("0", field.StartingNumber);
            Assert.Null(field.ListLevel);
            Assert.False(field.HasListName);
            Assert.Null(field.ListName);

            for (int i = 1; i < 4; i++)
            {
                field = (FieldListNum)doc.Range.Fields[i];

                TestUtil.VerifyField(FieldType.FieldListNum, " LISTNUM ", string.Empty, field);
                Assert.Null(field.StartingNumber);
                Assert.Null(field.ListLevel);
                Assert.False(field.HasListName);
                Assert.Null(field.ListName);
            }

            field = (FieldListNum)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldListNum, " LISTNUM  \\l 2", string.Empty, field);
            Assert.Null(field.StartingNumber);
            Assert.AreEqual("2", field.ListLevel);
            Assert.False(field.HasListName);
            Assert.Null(field.ListName);

            field = (FieldListNum)doc.Range.Fields[5];

            TestUtil.VerifyField(FieldType.FieldListNum, " LISTNUM  OutlineDefault \\s 1", string.Empty, field);
            Assert.AreEqual("1", field.StartingNumber);
            Assert.Null(field.ListLevel);
            Assert.IsTrue(field.HasListName);
            Assert.AreEqual("OutlineDefault", field.ListName);
        }

        [Test]
        public void MergeField()
        {
            //ExStart
            //ExFor:FieldMergeField
            //ExFor:FieldMergeField.FieldName
            //ExFor:FieldMergeField.FieldNameNoPrefix
            //ExFor:FieldMergeField.IsMapped
            //ExFor:FieldMergeField.IsVerticalFormatting
            //ExFor:FieldMergeField.TextAfter
            //ExSummary:Shows how to use MERGEFIELD fields to perform a mail merge.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a data table to be used as a mail merge data source.
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Courtesy Title");
            table.Columns.Add("First Name");
            table.Columns.Add("Last Name");
            table.Rows.Add("Mr.", "John", "Doe");
            table.Rows.Add("Mrs.", "Jane", "Cardholder");

            // Insert a MERGEFIELD with a FieldName attribute set to the name of a column in the data source.
            FieldMergeField fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Courtesy Title";
            fieldMergeField.IsMapped = true;
            fieldMergeField.IsVerticalFormatting = false;

            // We can apply text before and after the value that this field accepts when the merge takes place.
            fieldMergeField.TextBefore = "Dear ";
            fieldMergeField.TextAfter = " ";

            Assert.AreEqual(" MERGEFIELD  \"Courtesy Title\" \\m \\b \"Dear \" \\f \" \"", fieldMergeField.GetFieldCode());

            // Insert another MERGEFIELD for a different column in the data source.
            fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Last Name";
            fieldMergeField.TextAfter = ":";

            doc.UpdateFields();
            doc.MailMerge.Execute(table);

            Assert.AreEqual("Dear Mr. Doe:\u000cDear Mrs. Cardholder:", doc.GetText().Trim());
            //ExEnd

            Assert.That(doc.Range.Fields, Is.Empty);
        }

        //ExStart
        //ExFor:FieldToc
        //ExFor:FieldToc.BookmarkName
        //ExFor:FieldToc.CustomStyles
        //ExFor:FieldToc.EntrySeparator
        //ExFor:FieldToc.HeadingLevelRange
        //ExFor:FieldToc.HideInWebLayout
        //ExFor:FieldToc.InsertHyperlinks
        //ExFor:FieldToc.PageNumberOmittingLevelRange
        //ExFor:FieldToc.PreserveLineBreaks
        //ExFor:FieldToc.PreserveTabs
        //ExFor:FieldToc.UpdatePageNumbers
        //ExFor:FieldToc.UseParagraphOutlineLevel
        //ExFor:FieldOptions.CustomTocStyleSeparator
        //ExSummary:Shows how to insert a TOC, and populate it with entries based on heading styles.
        [Test] //ExSkip
        public void FieldToc()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("MyBookmark");

            // Insert a TOC field, which will compile all headings into a table of contents.
            // For each heading, this field will create a line with the text in that heading style to the left,
            // and the page the heading appears on to the right.
            FieldToc field = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

            // Use the BookmarkName attribute to only list headings
            // that appear within the bounds of a bookmark with the "MyBookmark" name.
            field.BookmarkName = "MyBookmark";

            // Text with a built-in heading style, such as "Heading 1", applied to it will count as a heading.
            // We can name additional styles to be picked up as headings by the TOC in this attribute,
            // as well as their TOC levels.
            field.CustomStyles = "Quote; 6; Intense Quote; 7";

            // By default, Styles/TOC levels are separated in the CustomStyles attribute by a comma,
            // but we can set a custom delimiter in this attribute.
            doc.FieldOptions.CustomTocStyleSeparator = ";";

            // Configure the field to exclude any headings that have TOC levels outside of this range.
            field.HeadingLevelRange = "1-3";

            // The TOC will not display the page numbers of headings whose TOC levels are within this range.
            field.PageNumberOmittingLevelRange = "2-5";

            // Set a custom string that will be placed between every heading and its page number.
            field.EntrySeparator = "-";
            field.InsertHyperlinks = true;
            field.HideInWebLayout = false;
            field.PreserveLineBreaks = true;
            field.PreserveTabs = true;
            field.UseParagraphOutlineLevel = false;

            InsertNewPageWithHeading(builder, "First entry", "Heading 1");
            builder.Writeln("Paragraph text.");
            InsertNewPageWithHeading(builder, "Second entry", "Heading 1");
            InsertNewPageWithHeading(builder, "Third entry", "Quote");
            InsertNewPageWithHeading(builder, "Fourth entry", "Intense Quote");

            // These two headings will have the page numbers omitted because they are within the "2-5" range.
            InsertNewPageWithHeading(builder, "Fifth entry", "Heading 2");
            InsertNewPageWithHeading(builder, "Sixth entry", "Heading 3");

            // This entry will be omitted because "Heading 4" is outside of the "1-3" range that we have set earlier.
            InsertNewPageWithHeading(builder, "Seventh entry", "Heading 4");

            builder.EndBookmark("MyBookmark");
            builder.Writeln("Paragraph text.");

            // This entry will be omitted because it is outside the bookmark specified by the TOC.
            InsertNewPageWithHeading(builder, "Eighth entry", "Heading 1");

            Assert.AreEqual(" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field.GetFieldCode());

            field.UpdatePageNumbers();
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.TOC.docx");
            TestFieldToc(doc); //ExSkip
        }

        /// <summary>
        /// Start a new page and insert a paragraph of a specified style.
        /// </summary>
        public void InsertNewPageWithHeading(DocumentBuilder builder, string captionText, string styleName)
        {
            builder.InsertBreak(BreakType.PageBreak);
            string originalStyle = builder.ParagraphFormat.StyleName;
            builder.ParagraphFormat.Style = builder.Document.Styles[styleName];
            builder.Writeln(captionText);
            builder.ParagraphFormat.Style = builder.Document.Styles[originalStyle];
        }
        //ExEnd

        private void TestFieldToc(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);
            FieldToc field = (FieldToc)doc.Range.Fields[0];

            Assert.AreEqual("MyBookmark", field.BookmarkName);
            Assert.AreEqual("Quote; 6; Intense Quote; 7", field.CustomStyles);
            Assert.AreEqual("-", field.EntrySeparator);
            Assert.AreEqual("1-3", field.HeadingLevelRange);
            Assert.AreEqual("2-5", field.PageNumberOmittingLevelRange);
            Assert.False(field.HideInWebLayout);
            Assert.True(field.InsertHyperlinks);
            Assert.True(field.PreserveLineBreaks);
            Assert.True(field.PreserveTabs);
            Assert.True(field.UpdatePageNumbers());
            Assert.False(field.UseParagraphOutlineLevel);
            Assert.AreEqual(" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field.GetFieldCode());
            Assert.AreEqual("\u0013 HYPERLINK \\l \"_Toc256000001\" \u0014First entry-\u0013 PAGEREF _Toc256000001 \\h \u00142\u0015\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000002\" \u0014Second entry-\u0013 PAGEREF _Toc256000002 \\h \u00143\u0015\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000003\" \u0014Third entry-\u0013 PAGEREF _Toc256000003 \\h \u00144\u0015\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000004\" \u0014Fourth entry-\u0013 PAGEREF _Toc256000004 \\h \u00145\u0015\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000005\" \u0014Fifth entry\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000006\" \u0014Sixth entry\u0015\r", field.Result);
        }

        //ExStart
        //ExFor:FieldToc.EntryIdentifier
        //ExFor:FieldToc.EntryLevelRange
        //ExFor:FieldTC
        //ExFor:FieldTC.OmitPageNumber
        //ExFor:FieldTC.Text
        //ExFor:FieldTC.TypeIdentifier
        //ExFor:FieldTC.EntryLevel
        //ExSummary:Shows how to insert a TOC field, and filter which TC fields end up as entries.
        [Test] //ExSkip
        public void FieldTocEntryIdentifier()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TOC field, which will compile all TC fields into a table of contents.
            FieldToc fieldToc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

            // Configure the field to only pick up TC entries of the "A" type,
            // and of an entry level between 1 and 3.
            fieldToc.EntryIdentifier = "A";
            fieldToc.EntryLevelRange = "1-3";

            Assert.AreEqual(" TOC  \\f A \\l 1-3", fieldToc.GetFieldCode());

            // These two entries will appear in the table.
            builder.InsertBreak(BreakType.PageBreak);
            InsertTocEntry(builder, "TC field 1", "A", "1");
            InsertTocEntry(builder, "TC field 2", "A", "2");

            Assert.AreEqual(" TC  \"TC field 1\" \\n \\f A \\l 1", doc.Range.Fields[1].GetFieldCode());

            // This entry will be omitted from the table because it has a type that is different from "A".
            InsertTocEntry(builder, "TC field 3", "B", "1");

            // This entry will be omitted from the table because it has an entry level outside of the 1-3 range.
            InsertTocEntry(builder, "TC field 4", "A", "5");
            
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.TC.docx");
            TestFieldTocEntryIdentifier(doc); //ExSkip
        }

        /// <summary>
        /// Use a document builder to insert a TC field.
        /// </summary>
        public void InsertTocEntry(DocumentBuilder builder, string text, string typeIdentifier, string entryLevel)
        {
            FieldTC fieldTc = (FieldTC)builder.InsertField(FieldType.FieldTOCEntry, true);
            fieldTc.OmitPageNumber = true;
            fieldTc.Text = text;
            fieldTc.TypeIdentifier = typeIdentifier;
            fieldTc.EntryLevel = entryLevel;
        }
        //ExEnd

        private void TestFieldTocEntryIdentifier(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);
            FieldToc fieldToc = (FieldToc)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldTOC, " TOC  \\f A \\l 1-3", "TC field 1\rTC field 2\r", fieldToc);
            Assert.AreEqual("A", fieldToc.EntryIdentifier);
            Assert.AreEqual("1-3", fieldToc.EntryLevelRange);

            FieldTC fieldTc = (FieldTC)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldTOCEntry, " TC  \"TC field 1\" \\n \\f A \\l 1", string.Empty, fieldTc);
            Assert.True(fieldTc.OmitPageNumber);
            Assert.AreEqual("TC field 1", fieldTc.Text);
            Assert.AreEqual("A", fieldTc.TypeIdentifier);
            Assert.AreEqual("1", fieldTc.EntryLevel);

            fieldTc = (FieldTC)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldTOCEntry, " TC  \"TC field 2\" \\n \\f A \\l 2", string.Empty, fieldTc);
            Assert.True(fieldTc.OmitPageNumber);
            Assert.AreEqual("TC field 2", fieldTc.Text);
            Assert.AreEqual("A", fieldTc.TypeIdentifier);
            Assert.AreEqual("2", fieldTc.EntryLevel);

            fieldTc = (FieldTC)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldTOCEntry, " TC  \"TC field 3\" \\n \\f B \\l 1", string.Empty, fieldTc);
            Assert.True(fieldTc.OmitPageNumber);
            Assert.AreEqual("TC field 3", fieldTc.Text);
            Assert.AreEqual("B", fieldTc.TypeIdentifier);
            Assert.AreEqual("1", fieldTc.EntryLevel);

            fieldTc = (FieldTC)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldTOCEntry, " TC  \"TC field 4\" \\n \\f A \\l 5", string.Empty, fieldTc);
            Assert.True(fieldTc.OmitPageNumber);
            Assert.AreEqual("TC field 4", fieldTc.Text);
            Assert.AreEqual("A", fieldTc.TypeIdentifier);
            Assert.AreEqual("5", fieldTc.EntryLevel);
        }

        [Test]
        public void TocSeqPrefix()
        {
            //ExStart
            //ExFor:FieldToc
            //ExFor:FieldToc.TableOfFiguresLabel
            //ExFor:FieldToc.PrefixedSequenceIdentifier
            //ExFor:FieldToc.SequenceSeparator
            //ExFor:FieldSeq
            //ExFor:FieldSeq.SequenceIdentifier
            //ExSummary:Shows how to populate a TOC field with entries using SEQ fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
            // Each entry contains the paragraph that contains the SEQ field,
            // and the number of the page that the field appears on.
            FieldToc fieldToc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

            // SEQ fields display a count that increments at each SEQ field.
            // These fields also maintain separate counts for each unique named sequence
            // identified by the SEQ field's "SequenceIdentifier" attribute.
            // Use the "TableOfFiguresLabel" attribute to name a main sequence for the TOC.
            // Now, this TOC will only create entries out of SEQ fields
            // that have their "SequenceIdentifier" set to "MySequence".
            fieldToc.TableOfFiguresLabel = "MySequence";

            // We can name another SEQ field sequence in the "PrefixedSequenceIdentifier" attribute.
            // SEQ fields from this prefix sequence will not create TOC entries. 
            // Every TOC entry created from a main sequence SEQ field
            // will now also display the count that the prefix sequence is currently on
            // at the location of main sequence SEQ field that created the entry.
            fieldToc.PrefixedSequenceIdentifier = "PrefixSequence";

            // Each TOC entry will display the prefix sequence count immediately to the left
            // of the page number that the main sequence SEQ field appears on.
            // We can specify a custom separator that will appear between these two numbers.
            fieldToc.SequenceSeparator = ">";

            Assert.AreEqual(" TOC  \\c MySequence \\s PrefixSequence \\d >", fieldToc.GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);

            // There are two ways of using SEQ fields to populate this TOC.
            // 1 -  Inserting a SEQ field that belongs to the TOC's prefix sequence:
            // This field will increment the SEQ sequence count for the "PrefixSequence" by 1.
            // Since this field does not belong to the main sequence identified
            // by the "TableOfFiguresLabel" attribute of the TOC, it will not show up as an entry.
            FieldSeq fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "PrefixSequence";
            builder.InsertParagraph();

            Assert.AreEqual(" SEQ  PrefixSequence", fieldSeq.GetFieldCode());

            // 2 -  Inserting a SEQ field that belongs to the TOC's main sequence:
            // This SEQ field will create an entry in the TOC.
            // The TOC entry will contain the paragraph that the SEQ field is in,
            // as well as the number of the page that it appears on.
            // This entry will also display the count that the prefix sequence is currently at,
            // separated from the page number by the value in the TOC's SeqenceSeparator attribute.
            // The "PrefixSequence" count is at 1, this main sequence SEQ field is on page 2,
            // and the separator is ">", so entry will display "1>2".
            builder.Write("First TOC entry, MySequence #");
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";

            Assert.AreEqual(" SEQ  MySequence", fieldSeq.GetFieldCode());

            // Insert a page, advance the prefix sequence by 2, and insert a SEQ field which will create a TOC entry afterwards.
            // The prefix sequence is now at 2, and the main sequence SEQ field is on page 3,
            // so the TOC entry will display "2>3" at its page count.
            builder.InsertBreak(BreakType.PageBreak);
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "PrefixSequence";
            builder.InsertParagraph();
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            builder.Write("Second TOC entry, MySequence #");
            fieldSeq.SequenceIdentifier = "MySequence";

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.TOC.SEQ.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.TOC.SEQ.docx");

            Assert.AreEqual(9, doc.Range.Fields.Count);

            fieldToc = (FieldToc)doc.Range.Fields[0];
            Console.WriteLine(fieldToc.DisplayResult);
            TestUtil.VerifyField(FieldType.FieldTOC, " TOC  \\c MySequence \\s PrefixSequence \\d >",
                "First TOC entry, MySequence #12\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r2" +
                "Second TOC entry, MySequence #\t\u0013 SEQ PrefixSequence _Toc256000001 \\* ARABIC \u00142\u0015>\u0013 PAGEREF _Toc256000001 \\h \u00143\u0015\r", 
                fieldToc);
            Assert.AreEqual("MySequence", fieldToc.TableOfFiguresLabel);
            Assert.AreEqual("PrefixSequence", fieldToc.PrefixedSequenceIdentifier);
            Assert.AreEqual(">", fieldToc.SequenceSeparator);

            fieldSeq = (FieldSeq)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ PrefixSequence _Toc256000000 \\* ARABIC ", "1", fieldSeq);
            Assert.AreEqual("PrefixSequence", fieldSeq.SequenceIdentifier);

            // Byproduct field created by Aspose.Words
            FieldPageRef fieldPageRef = (FieldPageRef)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldPageRef, " PAGEREF _Toc256000000 \\h ", "2", fieldPageRef);
            Assert.AreEqual("PrefixSequence", fieldSeq.SequenceIdentifier);
            Assert.AreEqual("_Toc256000000", fieldPageRef.BookmarkName);

            fieldSeq = (FieldSeq)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ PrefixSequence _Toc256000001 \\* ARABIC ", "2", fieldSeq);
            Assert.AreEqual("PrefixSequence", fieldSeq.SequenceIdentifier);

            fieldPageRef = (FieldPageRef)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldPageRef, " PAGEREF _Toc256000001 \\h ", "3", fieldPageRef);
            Assert.AreEqual("PrefixSequence", fieldSeq.SequenceIdentifier);
            Assert.AreEqual("_Toc256000001", fieldPageRef.BookmarkName);

            fieldSeq = (FieldSeq)doc.Range.Fields[5];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  PrefixSequence", "1", fieldSeq);
            Assert.AreEqual("PrefixSequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[6];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence", "1", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[7];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  PrefixSequence", "2", fieldSeq);
            Assert.AreEqual("PrefixSequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[8];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence", "2", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);
        }

        [Test]
        public void TocSeqNumbering()
        {
            //ExStart
            //ExFor:FieldSeq
            //ExFor:FieldSeq.InsertNextNumber
            //ExFor:FieldSeq.ResetHeadingLevel
            //ExFor:FieldSeq.ResetNumber
            //ExFor:FieldSeq.SequenceIdentifier
            //ExSummary:Shows create numbering using SEQ fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // SEQ fields display a count that increments at each SEQ field.
            // These fields also maintain separate counts for each unique named sequence
            // identified by the SEQ field's "SequenceIdentifier" attribute.
            // Insert a SEQ field which will display the current count value of "MySequence",
            // after using the "ResetNumber" attribute to set it to 100.
            builder.Write("#");
            FieldSeq fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";
            fieldSeq.ResetNumber = "100";
            fieldSeq.Update();

            Assert.AreEqual(" SEQ  MySequence \\r 100", fieldSeq.GetFieldCode());
            Assert.AreEqual("100", fieldSeq.Result);

            // Display the next number in this sequence with another SEQ field.
            builder.Write(", #");
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";
            fieldSeq.Update();

            Assert.AreEqual("101", fieldSeq.Result);

            // Insert a level 1 heading.
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("This level 1 heading will reset MySequence to 1");
            builder.ParagraphFormat.Style = doc.Styles["Normal"];

            // Insert another SEQ field from the same sequence, and configure it
            // to reset the count at every heading with a level of 1.
            builder.Write("\n#");
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";
            fieldSeq.ResetHeadingLevel = "1";
            fieldSeq.Update();

            // The above heading is a level 1 heading, so the count for this sequence is reset to 1.
            Assert.AreEqual(" SEQ  MySequence \\s 1", fieldSeq.GetFieldCode());
            Assert.AreEqual("1", fieldSeq.Result);

            // Move to the next number of this sequence.
            builder.Write(", #");
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";
            fieldSeq.InsertNextNumber = true;
            fieldSeq.Update();

            Assert.AreEqual(" SEQ  MySequence \\n", fieldSeq.GetFieldCode());
            Assert.AreEqual("2", fieldSeq.Result);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.SEQ.ResetNumbering.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SEQ.ResetNumbering.docx");

            Assert.AreEqual(4, doc.Range.Fields.Count);

            fieldSeq = (FieldSeq)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence \\r 100", "100", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence", "101", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence \\s 1", "1", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence \\n", "2", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);
        }

        [Test]
        [Ignore("WORDSNET-18083")]
        public void TocSeqBookmark()
        {
            //ExStart
            //ExFor:FieldSeq
            //ExFor:FieldSeq.BookmarkName
            //ExSummary:Shows how to combine table of contents and sequence fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
            // Each entry contains the paragraph that contains the SEQ field,
            // and the number of the page that the field appears on.
            FieldToc fieldToc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

            // Configure this TOC field to only pick up SEQ fields that are within the bounds of a bookmark
            // named "TOCBookmark", and also have a SequenceIdentifier attribute with a value of "MySequence".
            fieldToc.TableOfFiguresLabel = "MySequence";
            fieldToc.BookmarkName = "TOCBookmark";
            builder.InsertBreak(BreakType.PageBreak);

            Assert.AreEqual(" TOC  \\c MySequence \\b TOCBookmark", fieldToc.GetFieldCode());

            // SEQ fields display a count that increments at each SEQ field.
            // These fields also maintain separate counts for each unique named sequence
            // identified by the SEQ field's "SequenceIdentifier" attribute.
            // Insert a SEQ field that has a sequence identifier that matches the TOC's
            // TableOfFiguresLabel attribute. This field will not create an entry in the TOC
            // since it is outside the bounds of a bookmark designated by "BookmarkName".
            builder.Write("MySequence #");
            FieldSeq fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";
            builder.Writeln(", will not show up in the TOC because it is outside of the bookmark.");

            builder.StartBookmark("TOCBookmark");

            // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" attribute, and is within the bounds of the bookmark.
            // The paragraph that contains this field will show up in the TOC as an entry.
            builder.Write("MySequence #");
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";
            builder.Writeln(", will show up in the TOC next to the entry for the above caption.");

            // This SEQ field's sequence does not match the TOC's "TableOfFiguresLabel" attribute,
            // and is within the bounds of the bookmark. Its paragraph will not show up in the TOC as an entry.
            builder.Write("MySequence #");
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "OtherSequence";
            builder.Writeln(", will not show up in the TOC because it's from a different sequence identifier.");

            // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" attribute, and is within the bounds of the bookmark.
            // This field also references another bookmark. The contents of that bookmark will appear in the TOC entry for this SEQ field.
            // The SEQ field itself will not display the contents of that bookmark.
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";
            fieldSeq.BookmarkName = "SEQBookmark";
            Assert.AreEqual(" SEQ  MySequence SEQBookmark", fieldSeq.GetFieldCode());

            // Create a bookmark with contents that will show up in the TOC entry due to the above SEQ field referencing it.
            builder.InsertBreak(BreakType.PageBreak);
            builder.StartBookmark("SEQBookmark");
            builder.Write("MySequence #");
            fieldSeq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            fieldSeq.SequenceIdentifier = "MySequence";
            builder.Writeln(", text from inside SEQBookmark.");
            builder.EndBookmark("SEQBookmark");

            builder.EndBookmark("TOCBookmark");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.SEQ.Bookmark.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SEQ.Bookmark.docx");

            Assert.AreEqual(8, doc.Range.Fields.Count);

            fieldToc = (FieldToc)doc.Range.Fields[0];
            string[] pageRefIds = fieldToc.Result.Split(' ').Where(s => s.StartsWith("_Toc")).ToArray();

            Assert.AreEqual(FieldType.FieldTOC, fieldToc.Type);
            Assert.AreEqual("MySequence", fieldToc.TableOfFiguresLabel);
            TestUtil.VerifyField(FieldType.FieldTOC, " TOC  \\c MySequence \\b TOCBookmark",
                $"MySequence #2, will show up in the TOC next to the entry for the above caption.\t\u0013 PAGEREF {pageRefIds[0]} \\h \u00142\u0015\r" +
                $"3MySequence #3, text from inside SEQBookmark.\t\u0013 PAGEREF {pageRefIds[1]} \\h \u00142\u0015\r", fieldToc);

            FieldPageRef fieldPageRef = (FieldPageRef)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldPageRef, $" PAGEREF {pageRefIds[0]} \\h ", "2", fieldPageRef);
            Assert.AreEqual(pageRefIds[0], fieldPageRef.BookmarkName);
            
            fieldPageRef = (FieldPageRef)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldPageRef, $" PAGEREF {pageRefIds[1]} \\h ", "2", fieldPageRef);
            Assert.AreEqual(pageRefIds[1], fieldPageRef.BookmarkName);

            fieldSeq = (FieldSeq)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence", "1", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence", "2", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[5];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  OtherSequence", "1", fieldSeq);
            Assert.AreEqual("OtherSequence", fieldSeq.SequenceIdentifier);

            fieldSeq = (FieldSeq)doc.Range.Fields[6];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence SEQBookmark", "3", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);
            Assert.AreEqual("SEQBookmark", fieldSeq.BookmarkName);

            fieldSeq = (FieldSeq)doc.Range.Fields[7];

            TestUtil.VerifyField(FieldType.FieldSequence, " SEQ  MySequence", "3", fieldSeq);
            Assert.AreEqual("MySequence", fieldSeq.SequenceIdentifier);
        }

        [Test]
        [Ignore("WORDSNET-13854")]
        public void FieldCitation()
        {
            //ExStart
            //ExFor:FieldCitation
            //ExFor:FieldCitation.AnotherSourceTag
            //ExFor:FieldCitation.FormatLanguageId
            //ExFor:FieldCitation.PageNumber
            //ExFor:FieldCitation.Prefix
            //ExFor:FieldCitation.SourceTag
            //ExFor:FieldCitation.Suffix
            //ExFor:FieldCitation.SuppressAuthor
            //ExFor:FieldCitation.SuppressTitle
            //ExFor:FieldCitation.SuppressYear
            //ExFor:FieldCitation.VolumeNumber
            //ExFor:FieldBibliography
            //ExFor:FieldBibliography.FormatLanguageId
            //ExSummary:Shows how to work with CITATION and BIBLIOGRAPHY fields.
            // Open a document that contains bibliographical sources
            // which we can find in Microsoft Word via References -> Citations & Bibliography -> Manage Sources.
            Document doc = new Document(MyDir + "Bibliography.docx");
            Assert.AreEqual(2, doc.Range.Fields.Count); //ExSkip

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Text to be cited with one source.");

            // Create a citation with just the page number, and the name of the author of the referenced book.
            FieldCitation fieldCitation = (FieldCitation)builder.InsertField(FieldType.FieldCitation, true);

            // We refer to sources using their tag names.
            fieldCitation.SourceTag = "Book1";
            fieldCitation.PageNumber = "85";
            fieldCitation.SuppressAuthor = false;
            fieldCitation.SuppressTitle = true;
            fieldCitation.SuppressYear = true;

            Assert.AreEqual(" CITATION  Book1 \\p 85 \\t \\y", fieldCitation.GetFieldCode());

            // Create a more detailed citation which cites two sources.
            builder.InsertParagraph();
            builder.Write("Text to be cited with two sources.");
            fieldCitation = (FieldCitation)builder.InsertField(FieldType.FieldCitation, true);
            fieldCitation.SourceTag = "Book1";
            fieldCitation.AnotherSourceTag = "Book2";
            fieldCitation.FormatLanguageId = "en-US";
            fieldCitation.PageNumber = "19";
            fieldCitation.Prefix = "Prefix ";
            fieldCitation.Suffix = " Suffix";
            fieldCitation.SuppressAuthor = false;
            fieldCitation.SuppressTitle = false;
            fieldCitation.SuppressYear = false;
            fieldCitation.VolumeNumber = "VII";

            Assert.AreEqual(" CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII", fieldCitation.GetFieldCode());

            // We can use a BIBLIOGRAPHY field to display all the sources within the document.
            builder.InsertBreak(BreakType.PageBreak);
            FieldBibliography fieldBibliography = (FieldBibliography)builder.InsertField(FieldType.FieldBibliography, true);
            fieldBibliography.FormatLanguageId = "1124";

            Assert.AreEqual(" BIBLIOGRAPHY  \\l 1124", fieldBibliography.GetFieldCode());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.CITATION.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.CITATION.docx");

            Assert.AreEqual(5, doc.Range.Fields.Count);

            fieldCitation = (FieldCitation)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldCitation, " CITATION  Book1 \\p 85 \\t \\y", " (Doe, p. 85)", fieldCitation);
            Assert.AreEqual("Book1", fieldCitation.SourceTag);
            Assert.AreEqual("85", fieldCitation.PageNumber);
            Assert.False(fieldCitation.SuppressAuthor);
            Assert.True(fieldCitation.SuppressTitle);
            Assert.True(fieldCitation.SuppressYear);

            fieldCitation = (FieldCitation)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldCitation, 
                " CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII", 
                " (Doe, 2018; Prefix Cardholder, 2018, VII:19 Suffix)", fieldCitation);
            Assert.AreEqual("Book1", fieldCitation.SourceTag);
            Assert.AreEqual("Book2", fieldCitation.AnotherSourceTag);
            Assert.AreEqual("en-US", fieldCitation.FormatLanguageId);
            Assert.AreEqual("Prefix ", fieldCitation.Prefix);
            Assert.AreEqual(" Suffix", fieldCitation.Suffix);
            Assert.AreEqual("19", fieldCitation.PageNumber);
            Assert.False(fieldCitation.SuppressAuthor);
            Assert.False(fieldCitation.SuppressTitle);
            Assert.False(fieldCitation.SuppressYear);
            Assert.AreEqual("VII", fieldCitation.VolumeNumber);

            fieldBibliography = (FieldBibliography)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldBibliography, " BIBLIOGRAPHY  \\l 1124",
                "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography);
            Assert.AreEqual("1124", fieldBibliography.FormatLanguageId);

            fieldCitation = (FieldCitation)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldCitation, " CITATION Book1 \\l 1033 ", "(Doe, 2018)", fieldCitation);
            Assert.AreEqual("Book1", fieldCitation.SourceTag);
            Assert.AreEqual("1033", fieldCitation.FormatLanguageId);

            fieldBibliography = (FieldBibliography)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldBibliography, " BIBLIOGRAPHY ", 
                "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography);
        }

        [Test]
        public void FieldData()
        {
            //ExStart
            //ExFor:FieldData
            //ExSummary:Shows how to insert a DATA field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldData field = (FieldData)builder.InsertField(FieldType.FieldData, true);
            Assert.AreEqual(" DATA ", field.GetFieldCode());
            //ExEnd
            
            TestUtil.VerifyField(FieldType.FieldData, " DATA ", string.Empty, DocumentHelper.SaveOpen(doc).Range.Fields[0]);
        }

        [Test]
        public void FieldInclude()
        {
            //ExStart
            //ExFor:FieldInclude
            //ExFor:FieldInclude.BookmarkName
            //ExFor:FieldInclude.LockFields
            //ExFor:FieldInclude.SourceFullName
            //ExFor:FieldInclude.TextConverter
            //ExSummary:Shows how to create an INCLUDE field, and set its properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We can use an INCLUDE field to import a portion of another document in the local file system.
            // The bookmark from the other document that we reference with this field contains this imported portion.
            FieldInclude field = (FieldInclude)builder.InsertField(FieldType.FieldInclude, true);
            field.SourceFullName = MyDir + "Bookmarks.docx";
            field.BookmarkName = "MyBookmark1";
            field.LockFields = false;
            field.TextConverter = "Microsoft Word";

            Assert.True(Regex.Match(field.GetFieldCode(), " INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"").Success);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INCLUDE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.INCLUDE.docx");
            field = (FieldInclude)doc.Range.Fields[0];

            Assert.AreEqual(FieldType.FieldInclude, field.Type);
            Assert.AreEqual("First bookmark.", field.Result);
            Assert.True(Regex.Match(field.GetFieldCode(), " INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"").Success);

            Assert.AreEqual(MyDir + "Bookmarks.docx", field.SourceFullName);
            Assert.AreEqual("MyBookmark1", field.BookmarkName);
            Assert.False(field.LockFields);
            Assert.AreEqual("Microsoft Word", field.TextConverter);
        }

        [Test]
        public void FieldIncludePicture()
        {
            //ExStart
            //ExFor:FieldIncludePicture
            //ExFor:FieldIncludePicture.GraphicFilter
            //ExFor:FieldIncludePicture.IsLinked
            //ExFor:FieldIncludePicture.ResizeHorizontally
            //ExFor:FieldIncludePicture.ResizeVertically
            //ExFor:FieldIncludePicture.SourceFullName
            //ExFor:FieldImport
            //ExFor:FieldImport.GraphicFilter
            //ExFor:FieldImport.IsLinked
            //ExFor:FieldImport.SourceFullName
            //ExSummary:Shows how to insert images using IMPORT and INCLUDEPICTURE fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two similar field types which we can use to display images linked from the local file system.
            // 1 -  The INCLUDEPICTURE field:
            FieldIncludePicture fieldIncludePicture = (FieldIncludePicture)builder.InsertField(FieldType.FieldIncludePicture, true);
            fieldIncludePicture.SourceFullName = ImageDir + "Transparent background logo.png";

            Assert.True(Regex.Match(fieldIncludePicture.GetFieldCode(), " INCLUDEPICTURE  .*").Success);

            // Apply the PNG32.FLT filter.
            fieldIncludePicture.GraphicFilter = "PNG32";
            fieldIncludePicture.IsLinked = true;
            fieldIncludePicture.ResizeHorizontally = true;
            fieldIncludePicture.ResizeVertically = true;

            // 2 -  The IMPORT field:
            FieldImport fieldImport = (FieldImport)builder.InsertField(FieldType.FieldImport, true);
            fieldImport.SourceFullName = ImageDir + "Transparent background logo.png";
            fieldImport.GraphicFilter = "PNG32";
            fieldImport.IsLinked = true;

            Assert.True(Regex.Match(fieldImport.GetFieldCode(), " IMPORT  .* \\\\c PNG32 \\\\d").Success);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.IMPORT.INCLUDEPICTURE.docx");
            //ExEnd

            Assert.AreEqual(ImageDir + "Transparent background logo.png", fieldIncludePicture.SourceFullName);
            Assert.AreEqual("PNG32", fieldIncludePicture.GraphicFilter);
            Assert.True(fieldIncludePicture.IsLinked);
            Assert.True(fieldIncludePicture.ResizeHorizontally);
            Assert.True(fieldIncludePicture.ResizeVertically);

            Assert.AreEqual(ImageDir + "Transparent background logo.png", fieldImport.SourceFullName);
            Assert.AreEqual("PNG32", fieldImport.GraphicFilter);
            Assert.True(fieldImport.IsLinked);
            
            doc = new Document(ArtifactsDir + "Field.IMPORT.INCLUDEPICTURE.docx");

            // The INCLUDEPICTURE fields have been converted into shapes with linked images during loading
            Assert.AreEqual(0, doc.Range.Fields.Count);
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Shape, true).Count);

            Shape image = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.True(image.IsImage);
            Assert.Null(image.ImageData.ImageBytes);
            Assert.AreEqual(ImageDir + "Transparent background logo.png", image.ImageData.SourceFullName);

            image = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.True(image.IsImage);
            Assert.Null(image.ImageData.ImageBytes);
            Assert.AreEqual(ImageDir + "Transparent background logo.png", image.ImageData.SourceFullName);
        }

        //ExStart
        //ExFor:FieldIncludeText
        //ExFor:FieldIncludeText.BookmarkName
        //ExFor:FieldIncludeText.Encoding
        //ExFor:FieldIncludeText.LockFields
        //ExFor:FieldIncludeText.MimeType
        //ExFor:FieldIncludeText.NamespaceMappings
        //ExFor:FieldIncludeText.SourceFullName
        //ExFor:FieldIncludeText.TextConverter
        //ExFor:FieldIncludeText.XPath
        //ExFor:FieldIncludeText.XslTransformation
        //ExSummary:Shows how to create an INCLUDETEXT field, and set its properties.
        [Test] //ExSkip
        [Ignore("WORDSNET-17543")] //ExSkip
        public void FieldIncludeText()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two ways that we can use INCLUDETEXT fields to display contents of an XML file in the local file system.
            // 1 -  Perform an XSL transformation on an XML document.
            FieldIncludeText fieldIncludeText = CreateFieldIncludeText(builder, MyDir + "CD collection data.xml", false, "text/xml", "XML", "ISO-8859-1");
            fieldIncludeText.XslTransformation = MyDir + "CD collection XSL transformation.xsl";

            builder.Writeln();

            // 2 -  Use an XPath to take specific elements from an XML document.
            fieldIncludeText = CreateFieldIncludeText(builder, MyDir + "CD collection data.xml", false, "text/xml", "XML", "ISO-8859-1");
            fieldIncludeText.NamespaceMappings = "xmlns:n='myNamespace'";
            fieldIncludeText.XPath = "/catalog/cd/title";

            doc.Save(ArtifactsDir + "Field.INCLUDETEXT.docx");
            TestFieldIncludeText(new Document(ArtifactsDir + "Field.INCLUDETEXT.docx")); //ExSkip
        }

        /// <summary>
        /// Use a document builder to insert an INCLUDETEXT field with custom properties.
        /// </summary>
        public FieldIncludeText CreateFieldIncludeText(DocumentBuilder builder, string sourceFullName, bool lockFields, string mimeType, string textConverter, string encoding)
        {
            FieldIncludeText fieldIncludeText = (FieldIncludeText)builder.InsertField(FieldType.FieldIncludeText, true);
            fieldIncludeText.SourceFullName = sourceFullName;
            fieldIncludeText.LockFields = lockFields;
            fieldIncludeText.MimeType = mimeType;
            fieldIncludeText.TextConverter = textConverter;
            fieldIncludeText.Encoding = encoding;

            return fieldIncludeText;
        }
        //ExEnd

        private void TestFieldIncludeText(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);

            FieldIncludeText fieldIncludeText = (FieldIncludeText)doc.Range.Fields[0];
            Assert.AreEqual(MyDir + "CD collection data.xml", fieldIncludeText.SourceFullName);
            Assert.AreEqual(MyDir + "CD collection XSL transformation.xsl", fieldIncludeText.XslTransformation);
            Assert.False(fieldIncludeText.LockFields);
            Assert.AreEqual("text/xml", fieldIncludeText.MimeType);
            Assert.AreEqual("XML", fieldIncludeText.TextConverter);
            Assert.AreEqual("ISO-8859-1", fieldIncludeText.Encoding);
            Assert.AreEqual(" INCLUDETEXT  \"" + MyDir.Replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\t \"" + 
                            MyDir.Replace("\\", "\\\\") + "CD collection XSL transformation.xsl\"", 
                fieldIncludeText.GetFieldCode());
            Assert.True(fieldIncludeText.Result.StartsWith("My CD Collection"));

            XmlDocument cdCollectionData = new XmlDocument();
            cdCollectionData.LoadXml(File.ReadAllText(MyDir + "CD collection data.xml"));
            XmlNode catalogData = cdCollectionData.ChildNodes[0];

            XmlDocument cdCollectionXslTransformation = new XmlDocument();
            cdCollectionXslTransformation.LoadXml(File.ReadAllText(MyDir + "CD collection XSL transformation.xsl"));

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            XmlNamespaceManager manager = new XmlNamespaceManager(cdCollectionXslTransformation.NameTable);
            manager.AddNamespace("xsl", "http://www.w3.org/1999/XSL/Transform");

            for (int i = 0; i < table.Rows.Count; i++)
                for (int j = 0; j < table.Rows[i].Count; j++)
                {
                    if (i == 0)
                    {
                        // When on the first row from the input document's table, ensure that all table's cells match all XML element Names.
                        for (int k = 0; k < table.Rows.Count - 1; k++)
                            Assert.AreEqual(catalogData.ChildNodes[k].ChildNodes[j].Name,
                                table.Rows[i].Cells[j].GetText().Replace(ControlChar.Cell, string.Empty).ToLower());

                        // Also make sure that the whole first row has the same color as the XSL transform.
                        Assert.AreEqual(cdCollectionXslTransformation.SelectNodes("//xsl:stylesheet/xsl:template/html/body/table/tr", manager)[0].Attributes.GetNamedItem("bgcolor").Value,
                            ColorTranslator.ToHtml(table.Rows[i].Cells[j].CellFormat.Shading.BackgroundPatternColor).ToLower());
                    }
                    else
                    {
                        // When on all other rows of the input document's table, ensure that cell contents match XML element Values.
                        Assert.AreEqual(catalogData.ChildNodes[i - 1].ChildNodes[j].FirstChild.Value,
                            table.Rows[i].Cells[j].GetText().Replace(ControlChar.Cell, string.Empty));
                        Assert.AreEqual(Color.Empty, table.Rows[i].Cells[j].CellFormat.Shading.BackgroundPatternColor);
                    }

                    Assert.AreEqual(
                        double.Parse(cdCollectionXslTransformation.SelectNodes("//xsl:stylesheet/xsl:template/html/body/table", manager)[0].Attributes.GetNamedItem("border").Value) * 0.75, 
                        table.FirstRow.RowFormat.Borders.Bottom.LineWidth);
                }

            fieldIncludeText = (FieldIncludeText)doc.Range.Fields[1];
            Assert.AreEqual(MyDir + "CD collection data.xml", fieldIncludeText.SourceFullName);
            Assert.Null(fieldIncludeText.XslTransformation);
            Assert.False(fieldIncludeText.LockFields);
            Assert.AreEqual("text/xml", fieldIncludeText.MimeType);
            Assert.AreEqual("XML", fieldIncludeText.TextConverter);
            Assert.AreEqual("ISO-8859-1", fieldIncludeText.Encoding);
            Assert.AreEqual(" INCLUDETEXT  \"" + MyDir.Replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\n xmlns:n='myNamespace' \\x /catalog/cd/title", 
                fieldIncludeText.GetFieldCode());

            string expectedFieldResult = "";
            for (int i = 0; i < catalogData.ChildNodes.Count; i++)
            {
                expectedFieldResult += catalogData.ChildNodes[i].ChildNodes[0].ChildNodes[0].Value;
            }

            Assert.AreEqual(expectedFieldResult, fieldIncludeText.Result);
        }

        [Test]
        [Ignore("WORDSNET-17545")]
        public void FieldHyperlink()
        {
            //ExStart
            //ExFor:FieldHyperlink
            //ExFor:FieldHyperlink.Address
            //ExFor:FieldHyperlink.IsImageMap
            //ExFor:FieldHyperlink.OpenInNewWindow
            //ExFor:FieldHyperlink.ScreenTip
            //ExFor:FieldHyperlink.SubAddress
            //ExFor:FieldHyperlink.Target
            //ExSummary:Shows how to use HYPERLINK fields to link to documents in the local file system.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldHyperlink field = (FieldHyperlink)builder.InsertField(FieldType.FieldHyperlink, true);

            // When we click this HYPERLINK field in Microsoft Word,
            // it will open the linked document, and also place the cursor at the specified bookmark.
            field.Address = MyDir + "Bookmarks.docx";
            field.SubAddress = "MyBookmark3";
            field.ScreenTip = "Open " + field.Address + " on bookmark " + field.SubAddress + " in a new window";

            builder.Writeln();

            // When we click this HYPERLINK field in Microsoft Word,
            // it will open the linked document, and automatically scroll down to the specified iframe.
            field = (FieldHyperlink)builder.InsertField(FieldType.FieldHyperlink, true);
            field.Address = MyDir + "Iframes.html";
            field.ScreenTip = "Open " + field.Address;
            field.Target = "iframe_3";
            field.OpenInNewWindow = true;
            field.IsImageMap = false;

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.HYPERLINK.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.HYPERLINK.docx");
            field = (FieldHyperlink)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldHyperlink, 
                " HYPERLINK \"" + MyDir.Replace("\\", "\\\\") + "Bookmarks.docx\" \\l \"MyBookmark3\" \\o \"Open " + MyDir + "Bookmarks.docx on bookmark MyBookmark3 in a new window\" ",
                MyDir + "Bookmarks.docx - MyBookmark3", field);
            Assert.AreEqual(MyDir + "Bookmarks.docx", field.Address);
            Assert.AreEqual("MyBookmark3", field.SubAddress);
            Assert.AreEqual("Open " + field.Address.Replace("\\", string.Empty) + " on bookmark " + field.SubAddress + " in a new window", field.ScreenTip);

            field = (FieldHyperlink)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldHyperlink, " HYPERLINK \"file:///" + MyDir.Replace("\\", "\\\\").Replace(" ", "%20") + "Iframes.html\" \\t \"iframe_3\" \\o \"Open " + MyDir.Replace("\\", "\\\\") + "Iframes.html\" ",
                MyDir + "Iframes.html", field);
            Assert.AreEqual("file:///" + MyDir.Replace(" ", "%20") + "Iframes.html", field.Address);
            Assert.AreEqual("Open " + MyDir + "Iframes.html", field.ScreenTip);
            Assert.AreEqual("iframe_3", field.Target);
            Assert.False(field.OpenInNewWindow);
            Assert.False(field.IsImageMap);
        }

        //ExStart
        //ExFor:MergeFieldImageDimension
        //ExFor:MergeFieldImageDimension.#ctor
        //ExFor:MergeFieldImageDimension.#ctor(Double)
        //ExFor:MergeFieldImageDimension.#ctor(Double,MergeFieldImageDimensionUnit)
        //ExFor:MergeFieldImageDimension.Unit
        //ExFor:MergeFieldImageDimension.Value
        //ExFor:MergeFieldImageDimensionUnit
        //ExFor:ImageFieldMergingArgs
        //ExFor:ImageFieldMergingArgs.ImageFileName
        //ExFor:ImageFieldMergingArgs.ImageWidth
        //ExFor:ImageFieldMergingArgs.ImageHeight
        //ExSummary:Shows how to set the dimensions of images as they are accepted by MERGEFIELDS during a mail merge.
        [Test] //ExSkip
        public void MergeFieldImageDimension()
        {
            Document doc = new Document();

            // Insert a MERGEFIELD which will accept images from a source during a mail merge. Use the field code to reference
            // a column in the data source which contains local system filenames of images we wish to use in the mail merge.
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldMergeField field = (FieldMergeField)builder.InsertField("MERGEFIELD Image:ImageColumn");

            // The data source should have such a column named "ImageColumn".
            Assert.AreEqual("Image:ImageColumn", field.FieldName);

            // Create a suitable data source.
            DataTable dataTable = new DataTable("Images");
            dataTable.Columns.Add(new DataColumn("ImageColumn"));
            dataTable.Rows.Add(ImageDir + "Logo.jpg");
            dataTable.Rows.Add(ImageDir + "Transparent background logo.png");
            dataTable.Rows.Add(ImageDir + "Enhanced Windows MetaFile.emf");
            
            // Configure a callback to modify the sizes of images at merge time, then execute the mail merge.
            doc.MailMerge.FieldMergingCallback = new MergedImageResizer(200, 200, MergeFieldImageDimensionUnit.Point);
            doc.MailMerge.Execute(dataTable);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.MERGEFIELD.ImageDimension.docx");
            TestMergeFieldImageDimension(doc); //ExSkip
        }

        /// <summary>
        /// Sets the size of all mail merged images to one defined width and height.
        /// </summary>
        private class MergedImageResizer : IFieldMergingCallback
        {
            public MergedImageResizer(double imageWidth, double imageHeight, MergeFieldImageDimensionUnit unit)
            {
                mImageWidth = imageWidth;
                mImageHeight = imageHeight;
                mUnit = unit;
            }

            public void FieldMerging(FieldMergingArgs e)
            {
                throw new NotImplementedException();
            }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {
                args.ImageFileName = args.FieldValue.ToString();
                args.ImageWidth = new MergeFieldImageDimension(mImageWidth, mUnit);
                args.ImageHeight = new MergeFieldImageDimension(mImageHeight, mUnit);

                Assert.AreEqual(mImageWidth, args.ImageWidth.Value);
                Assert.AreEqual(mUnit, args.ImageWidth.Unit);
                Assert.AreEqual(mImageHeight, args.ImageHeight.Value);
                Assert.AreEqual(mUnit, args.ImageHeight.Unit);
            }

            private readonly double mImageWidth;
            private readonly double mImageHeight;
            private readonly MergeFieldImageDimensionUnit mUnit;
        }
        //ExEnd

        private void TestMergeFieldImageDimension(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual(0, doc.Range.Fields.Count);
            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Shape, true).Count);

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, shape);
            Assert.AreEqual(200.0d, shape.Width);
            Assert.AreEqual(200.0d, shape.Height);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, shape);
            Assert.AreEqual(200.0d, shape.Width);
            Assert.AreEqual(200.0d, shape.Height);

            shape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            TestUtil.VerifyImageInShape(534, 534, ImageType.Emf, shape);
            Assert.AreEqual(200.0d, shape.Width);
            Assert.AreEqual(200.0d, shape.Height);
        }

        //ExStart
        //ExFor:ImageFieldMergingArgs.Image
        //ExSummary:Shows how to use a callback to customize image merging logic.
        [Test] //ExSkip
        public void MergeFieldImages()
        {
            Document doc = new Document();

            // Insert a MERGEFIELD which will accept images from a source during a mail merge. Use the field code to reference
            // a column in the data source which contains local system filenames of images we wish to use in the mail merge.
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldMergeField field = (FieldMergeField)builder.InsertField("MERGEFIELD Image:ImageColumn");

            // In this case, the field expects the data source to have such a column named "ImageColumn".
            Assert.AreEqual("Image:ImageColumn", field.FieldName);

            // Filenames can be lengthy, and if we can find a way to avoid storing them
            // in the data source, we may be able to considerably reduce its size.
            // Create a data source which refers to images using short names.
            DataTable dataTable = new DataTable("Images");
            dataTable.Columns.Add(new DataColumn("ImageColumn"));
            dataTable.Rows.Add("Dark logo");
            dataTable.Rows.Add("Transparent logo");

            // Assign a merging callback that contains all logic that processes those names,
            // and then execute the mail merge. 
            doc.MailMerge.FieldMergingCallback = new ImageFilenameCallback();
            doc.MailMerge.Execute(dataTable);

            doc.Save(ArtifactsDir + "Field.MERGEFIELD.Images.docx");
            TestMergeFieldImages(new Document(ArtifactsDir + "Field.MERGEFIELD.Images.docx")); //ExSkip
        }

        /// <summary>
        /// Contains a dictionary that maps names of images to local system filenames that contain these images.
        /// If a mail merge data source uses one of the names in the dictionary to refer to an image,
        /// this callback will pass the respective filename to the merge destination.
        /// </summary>
        private class ImageFilenameCallback : IFieldMergingCallback
        {
            public ImageFilenameCallback()
            {
                mImageFilenames = new Dictionary<string, string>();
                mImageFilenames.Add("Dark logo", ImageDir + "Logo.jpg");
                mImageFilenames.Add("Transparent logo", ImageDir + "Transparent background logo.png");
            }

            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                throw new NotImplementedException();
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                if (mImageFilenames.ContainsKey(args.FieldValue.ToString()))
                {
                    #if NET462 || JAVA
                    args.Image = Image.FromFile(mImageFilenames[args.FieldValue.ToString()]);
                    #elif NETCOREAPP2_1
                    args.Image = SKBitmap.Decode(mImageFilenames[args.FieldValue.ToString()]);
                    args.ImageFileName = mImageFilenames[args.FieldValue.ToString()];
                    #endif
                }
                
                Assert.NotNull(args.Image);
            }

            private readonly Dictionary<string, string> mImageFilenames;
        }
        //ExEnd

        private void TestMergeFieldImages(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual(0, doc.Range.Fields.Count);
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Shape, true).Count);

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, shape);
            Assert.AreEqual(300.0d, shape.Width);
            Assert.AreEqual(300.0d, shape.Height);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, shape);
            Assert.AreEqual(300.0d, shape.Width);
            Assert.AreEqual(300.0d, shape.Height);
        }

        [Test]
        [Ignore("WORDSNET-17524")]
        public void FieldIndexFilter()
        {
            //ExStart
            //ExFor:FieldIndex
            //ExFor:FieldIndex.BookmarkName
            //ExFor:FieldIndex.EntryType
            //ExFor:FieldXE
            //ExFor:FieldXE.EntryType
            //ExFor:FieldXE.Text
            //ExSummary:Shows how create an INDEX field, and then use XE fields to populate it with entries.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an INDEX field which will display an entry for each XE field found in the document.
            // Each entry will display the XE field's Text attribute value on the left side,
            // and the number of the page that contains the XE field on the right.
            // Multiple XE fields with matching Text attribute values are grouped into one INDEX field entry.
            FieldIndex index = (FieldIndex)builder.InsertField(FieldType.FieldIndex, true);

            // Configure the INDEX field to only display XE fields that are within the bounds
            // of a bookmark named "MainBookmark", and whose "EntryType" attributes have a value of "A".
            // For both INDEX and XE fields, the "EntryType" attribute only uses the first character of its string value.
            index.BookmarkName = "MainBookmark";
            index.EntryType = "A";

            Assert.AreEqual(" INDEX  \\b MainBookmark \\f A", index.GetFieldCode());

            // On a new page, start the bookmark with a name that matches the value
            // of the INDEX field's "BookmarkName" attribute.
            builder.InsertBreak(BreakType.PageBreak);
            builder.StartBookmark("MainBookmark");

            // This entry will be picked up by the INDEX field because it is inside the bookmark,
            // and its entry type also matches the INDEX field's entry type.
            FieldXE indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Index entry 1";
            indexEntry.EntryType = "A";

            Assert.AreEqual(" XE  \"Index entry 1\" \\f A", indexEntry.GetFieldCode());

            // Insert an XE field that will not appear in the INDEX because the entry types do not match.
            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Index entry 2";
            indexEntry.EntryType = "B";

            // End the bookmark and insert an XE field afterwards.
            // It is of the same type as the INDEX field, but will not appear
            // since it is outside the boundaries of the bookmark.
            builder.EndBookmark("MainBookmark");
            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Index entry 3";
            indexEntry.EntryType = "A";

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INDEX.XE.Filtering.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.INDEX.XE.Filtering.docx");
            index = (FieldIndex)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldIndex, " INDEX  \\b MainBookmark \\f A", "Index entry 1, 2\r", index);
            Assert.AreEqual("MainBookmark", index.BookmarkName);
            Assert.AreEqual("A", index.EntryType);

            indexEntry = (FieldXE)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  \"Index entry 1\" \\f A", string.Empty, indexEntry);
            Assert.AreEqual("Index entry 1", indexEntry.Text);
            Assert.AreEqual("A", indexEntry.EntryType);

            indexEntry = (FieldXE)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  \"Index entry 2\" \\f B", string.Empty, indexEntry);
            Assert.AreEqual("Index entry 2", indexEntry.Text);
            Assert.AreEqual("B", indexEntry.EntryType);

            indexEntry = (FieldXE)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  \"Index entry 3\" \\f A", string.Empty, indexEntry);
            Assert.AreEqual("Index entry 3", indexEntry.Text);
            Assert.AreEqual("A", indexEntry.EntryType);
        }

        [Test]
        [Ignore("WORDSNET-17524")]
        public void FieldIndexFormatting()
        {
            //ExStart
            //ExFor:FieldIndex
            //ExFor:FieldIndex.Heading
            //ExFor:FieldIndex.NumberOfColumns
            //ExFor:FieldIndex.LanguageId
            //ExFor:FieldIndex.LetterRange
            //ExFor:FieldXE
            //ExFor:FieldXE.IsBold
            //ExFor:FieldXE.IsItalic
            //ExFor:FieldXE.Text
            //ExSummary:Shows how to populate an INDEX field with entries using XE fields, and also modify its appearance.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an INDEX field which will display an entry for each XE field found in the document.
            // Each entry will display the XE field's Text attribute value on the left side,
            // and the number of the page that contains the XE field on the right.
            // Multiple XE fields with matching Text attribute values are grouped into one INDEX field entry.
            FieldIndex index = (FieldIndex)builder.InsertField(FieldType.FieldIndex, true);
            index.LanguageId = "1033";

            // Setting this attribute's value to "A" will group all the entries by their first letter,
            // and place that letter in uppercase above each group.
            index.Heading = "A";

            // Set the table created by the INDEX field to span over 2 columns.
            index.NumberOfColumns = "2";

            // Set any entries with starting letters outside the "a-c" character range to be omitted.
            index.LetterRange = "a-c";

            Assert.AreEqual(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.GetFieldCode());

            // These next two XE fields will show up under the "A" heading,
            // with their respective text stylings also applied to their page numbers .
            builder.InsertBreak(BreakType.PageBreak);
            FieldXE indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Apple";
            indexEntry.IsItalic = true;

            Assert.AreEqual(" XE  Apple \\i", indexEntry.GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Apricot";
            indexEntry.IsBold = true;

            Assert.AreEqual(" XE  Apricot \\b", indexEntry.GetFieldCode());

            // Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents.
            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Banana";

            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Cherry";

            // All INDEX field entries are sorted alphabetically, so this entry will show up under "A" with the other two.
            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Avocado";

            // This entry will be excluded because it starts with the letter "D",
            // which is outside the "a-c" character tange defined by the INDEX field's LetterRange attribute.
            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Durian";

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INDEX.XE.Formatting.docx");
            //ExEnd
            
            doc = new Document(ArtifactsDir + "Field.INDEX.XE.Formatting.docx");
            index = (FieldIndex)doc.Range.Fields[0];

            Assert.AreEqual("1033", index.LanguageId);
            Assert.AreEqual("A", index.Heading);
            Assert.AreEqual("2", index.NumberOfColumns);
            Assert.AreEqual("a-c", index.LetterRange);
            Assert.AreEqual(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.GetFieldCode());
            Assert.AreEqual("\fA\r" +
                            "Apple, 2\r" +
                            "Apricot, 3\r" +
                            "Avocado, 6\r" +
                            "B\r" +
                            "Banana, 4\r" +
                            "C\r" +
                            "Cherry, 5\r\f", index.Result);

            indexEntry = (FieldXE)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  Apple \\i", string.Empty, indexEntry);
            Assert.AreEqual("Apple", indexEntry.Text);
            Assert.False(indexEntry.IsBold);
            Assert.True(indexEntry.IsItalic);

            indexEntry = (FieldXE)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  Apricot \\b", string.Empty, indexEntry);
            Assert.AreEqual("Apricot", indexEntry.Text);
            Assert.True(indexEntry.IsBold);
            Assert.False(indexEntry.IsItalic);

            indexEntry = (FieldXE)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  Banana", string.Empty, indexEntry);
            Assert.AreEqual("Banana", indexEntry.Text);
            Assert.False(indexEntry.IsBold);
            Assert.False(indexEntry.IsItalic);

            indexEntry = (FieldXE)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  Cherry", string.Empty, indexEntry);
            Assert.AreEqual("Cherry", indexEntry.Text);
            Assert.False(indexEntry.IsBold);
            Assert.False(indexEntry.IsItalic);

            indexEntry = (FieldXE)doc.Range.Fields[5];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  Avocado", string.Empty, indexEntry);
            Assert.AreEqual("Avocado", indexEntry.Text);
            Assert.False(indexEntry.IsBold);
            Assert.False(indexEntry.IsItalic);

            indexEntry = (FieldXE)doc.Range.Fields[6];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  Durian", string.Empty, indexEntry);
            Assert.AreEqual("Durian", indexEntry.Text);
            Assert.False(indexEntry.IsBold);
            Assert.False(indexEntry.IsItalic);
        }

        [Test]
        [Ignore("WORDSNET-17524")]
        public void FieldIndexSequence()
        {
            //ExStart
            //ExFor:FieldIndex.HasSequenceName
            //ExFor:FieldIndex.SequenceName
            //ExFor:FieldIndex.SequenceSeparator
            //ExSummary:Shows how to split a document into portions by combining INDEX and SEQ fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an INDEX field which will display an entry for each XE field found in the document.
            // Each entry will display the XE field's Text attribute value on the left side,
            // and the number of the page that contains the XE field on the right.
            // Multiple XE fields with matching Text attribute values are grouped into one INDEX field entry.
            FieldIndex index = (FieldIndex)builder.InsertField(FieldType.FieldIndex, true);

            // In the SequenceName attribute, name a SEQ field sequence. Each entry of this INDEX field will now also display the
            // number that the sequence count is on at the location of the XE field that created this entry.
            index.SequenceName = "MySequence";

            // Set text that will around the sequence and page numbers in order to explain their meaning to the user.
            // An entry created with this configuration will display something like "MySequence at 1 on page 1" at its page number.
            // PageNumberSeparator and SequenceSeparator cannot be longer than 15 characters.
            index.PageNumberSeparator = "\tMySequence at ";
            index.SequenceSeparator = " on page ";
            Assert.IsTrue(index.HasSequenceName);

            Assert.AreEqual(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.GetFieldCode());

            // SEQ fields display a count that increments at each SEQ field.
            // These fields also maintain separate counts for each unique named sequence
            // identified by the SEQ field's "SequenceIdentifier" attribute.
            // Insert a SEQ field which moves the "MySequence" sequence to 1.
            // This field is treated as normal document text. It will not show up on an INDEX field's table of contents.
            builder.InsertBreak(BreakType.PageBreak);
            FieldSeq sequenceField = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            sequenceField.SequenceIdentifier = "MySequence";

            Assert.AreEqual(" SEQ  MySequence", sequenceField.GetFieldCode());

            // Insert an XE field which will create an entry in the INDEX field.
            // Since "MySequence" is at 1 and this XE field is on page 2, along with the custom separators we defined above,
            // this field's INDEX entry will display "Cat" on the left side, and "MySequence at 1 on page 2" on the right.
            FieldXE indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Cat";

            Assert.AreEqual(" XE  Cat", indexEntry.GetFieldCode());

            // Insert a page break, and use SEQ fields to advance "MySequence" to 3.
            builder.InsertBreak(BreakType.PageBreak);
            sequenceField = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            sequenceField.SequenceIdentifier = "MySequence";
            sequenceField = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            sequenceField.SequenceIdentifier = "MySequence";

            // Insert an XE field with the same Text attribute as the one above.
            // XE fields with matching Text attributes will be collected into one INDEX entry,
            // as opposed to each creating their own.
            // Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above.
            // The page number portion of that INDEX entry will now display "MySequence at 1 on page 2, 3 on page 3".
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Cat";

            // Insert an XE field with a new and unique Text attribute value.
            // This will add a new entry, with MySequence at 3 on page 4.
            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Dog";
            
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INDEX.XE.Sequence.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.INDEX.XE.Sequence.docx");
            index = (FieldIndex)doc.Range.Fields[0];

            Assert.AreEqual("MySequence", index.SequenceName);
            Assert.AreEqual("\tMySequence at ", index.PageNumberSeparator);
            Assert.AreEqual(" on page ", index.SequenceSeparator);
            Assert.True(index.HasSequenceName);
            Assert.AreEqual(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.GetFieldCode());
            Assert.AreEqual("Cat\tMySequence at 1 on page 2, 3 on page 3\r" +
                            "Dog\tMySequence at 3 on page 4\r", index.Result);

            Assert.AreEqual(3, doc.Range.Fields.Where(f => f.Type == FieldType.FieldSequence).Count());
        }

        [Test]
        [Ignore("WORDSNET-17524")]
        public void FieldIndexPageNumberSeparator()
        {
            //ExStart
            //ExFor:FieldIndex.HasPageNumberSeparator
            //ExFor:FieldIndex.PageNumberSeparator
            //ExFor:FieldIndex.PageNumberListSeparator
            //ExSummary:Shows how to edit the page number separator in an INDEX field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an INDEX field which will display an entry for each XE field found in the document.
            // Each entry will display the XE field's Text attribute value on the left side,
            // and the number of the page that contains the XE field on the right.
            // Multiple XE fields with matching Text attribute values are grouped into one INDEX field entry.
            FieldIndex index = (FieldIndex)builder.InsertField(FieldType.FieldIndex, true);

            // If our INDEX field has an entry for a group of XE fields,
            // the number of each page that contains each XE field will be displayed in the page number portion of the entry.
            // We can set custom separators to customize the appearance of these page numbers.
            index.PageNumberSeparator = ", on page(s) ";
            index.PageNumberListSeparator = " & ";
            
            Assert.AreEqual(" INDEX  \\e \", on page(s) \" \\l \" & \"", index.GetFieldCode());
            Assert.True(index.HasPageNumberSeparator);

            // After we insert these XE fields, the INDEX field will display "First entry, on page(s) 2 & 3 & 4".
            builder.InsertBreak(BreakType.PageBreak);
            FieldXE indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "First entry";

            Assert.AreEqual(" XE  \"First entry\"", indexEntry.GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "First entry";

            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "First entry";

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INDEX.XE.PageNumberList.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.INDEX.XE.PageNumberList.docx");
            index = (FieldIndex)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldIndex, " INDEX  \\e \", on page(s) \" \\l \" & \"", "First entry, on page(s) 2 & 3 & 4\r", index);
            Assert.AreEqual(", on page(s) ", index.PageNumberSeparator);
            Assert.AreEqual(" & ", index.PageNumberListSeparator);
            Assert.True(index.HasPageNumberSeparator);
        }

        [Test]
        [Ignore("WORDSNET-17524")]
        public void FieldIndexPageRangeBookmark()
        {
            //ExStart
            //ExFor:FieldIndex.PageRangeSeparator
            //ExFor:FieldXE.HasPageRangeBookmarkName
            //ExFor:FieldXE.PageRangeBookmarkName
            //ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an INDEX field which will display an entry for each XE field found in the document.
            // Each entry will display the XE field's Text attribute value on the left side,
            // and the number of the page that contains the XE field on the right.
            // Multiple XE fields with matching Text attribute values are grouped into one INDEX field entry.
            FieldIndex index = (FieldIndex)builder.InsertField(FieldType.FieldIndex, true);

            // For INDEX entries that display page ranges, we can specify a separator string
            // which will be placed between the number of the first page, and the number of the last.
            index.PageNumberSeparator = ", on page(s) ";
            index.PageRangeSeparator = " to ";

            Assert.AreEqual(" INDEX  \\e \", on page(s) \" \\g \" to \"", index.GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);
            FieldXE indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "My entry";

            // If an XE field names a bookmark using the PageRangeBookmarkName attribute,
            // its INDEX entry will show the range of pages that the bookmark spans
            // instead of the number of the page that contains the XE field.
            indexEntry.PageRangeBookmarkName = "MyBookmark";

            Assert.AreEqual(" XE  \"My entry\" \\r MyBookmark", indexEntry.GetFieldCode());
            Assert.True(indexEntry.HasPageRangeBookmarkName);

            // Insert a bookmark that starts on page 3, and ends on page 5.
            // The INDEX entry for the XE field that references this bookmark will display this page range.
            // In our table, the INDEX entry will display "My entry, on page(s) 3 to 5".
            builder.InsertBreak(BreakType.PageBreak);
            builder.StartBookmark("MyBookmark");
            builder.Write("Start of MyBookmark");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("End of MyBookmark");
            builder.EndBookmark("MyBookmark");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INDEX.XE.PageRangeBookmark.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.INDEX.XE.PageRangeBookmark.docx");
            index = (FieldIndex)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldIndex, " INDEX  \\e \", on page(s) \" \\g \" to \"", "My entry, on page(s) 3 to 5\r", index);
            Assert.AreEqual(", on page(s) ", index.PageNumberSeparator);
            Assert.AreEqual(" to ", index.PageRangeSeparator);

            indexEntry = (FieldXE)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  \"My entry\" \\r MyBookmark", string.Empty, indexEntry);
            Assert.AreEqual("My entry", indexEntry.Text);
            Assert.AreEqual("MyBookmark", indexEntry.PageRangeBookmarkName);
            Assert.True(indexEntry.HasPageRangeBookmarkName);
        }

        [Test]
        [Ignore("WORDSNET-17524")]
        public void FieldIndexCrossReferenceSeparator()
        {
            //ExStart
            //ExFor:FieldIndex.CrossReferenceSeparator
            //ExFor:FieldXE.PageNumberReplacement
            //ExSummary:Shows how to define cross references in an INDEX field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an INDEX field which will display an entry for each XE field found in the document.
            // Each entry will display the XE field's Text attribute value on the left side,
            // and the number of the page that contains the XE field on the right.
            // Multiple XE fields with matching Text attribute values are grouped into one INDEX field entry.
            FieldIndex index = (FieldIndex)builder.InsertField(FieldType.FieldIndex, true);

            // We can configure an XE field to get its INDEX entry to display a string instead of a page number.
            // First, for entries that substitute a page number with a string,
            // specify a custom separator that goes between the XE field's Text attribute value, and the string.
            index.CrossReferenceSeparator = ", see: ";

            Assert.AreEqual(" INDEX  \\k \", see: \"", index.GetFieldCode());

            // Insert an XE field which creates a regular INDEX entry which displays this field's page number,
            // and does not invoke the CrossReferenceSeparator value.
            // The entry for this XE field will display "Apple, 2".
            builder.InsertBreak(BreakType.PageBreak);
            FieldXE indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Apple";

            Assert.AreEqual(" XE  Apple", indexEntry.GetFieldCode());

            // Insert another XE field on page 3, and set a value for the PageNumberReplacement attribute.
            // This value will show up instead of the number of the page that this field is on,
            // and the INDEX field's CrossReferenceSeparator value will be placed in front of it.
            // The entry for this XE field will display "Banana, see: Tropical fruit".
            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Banana";
            indexEntry.PageNumberReplacement = "Tropical fruit";

            Assert.AreEqual(" XE  Banana \\t \"Tropical fruit\"", indexEntry.GetFieldCode());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INDEX.XE.CrossReferenceSeparator.docx");
            //ExEnd
            
            doc = new Document(ArtifactsDir + "Field.INDEX.XE.CrossReferenceSeparator.docx");
            index = (FieldIndex)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " INDEX  \\k \", see: \"",
                "Apple, 2\r" +
                "Banana, see: Tropical fruit\r", index);
            Assert.AreEqual(", see: ", index.CrossReferenceSeparator);

            indexEntry = (FieldXE)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  Apple", string.Empty, indexEntry);
            Assert.AreEqual("Apple", indexEntry.Text);
            Assert.Null(indexEntry.PageNumberReplacement);

            indexEntry = (FieldXE)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  Banana \\t \"Tropical fruit\"", string.Empty, indexEntry);
            Assert.AreEqual("Banana", indexEntry.Text);
            Assert.AreEqual("Tropical fruit", indexEntry.PageNumberReplacement);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Ignore("WORDSNET-17524")]
        public void FieldIndexSubheading(bool doRunSubentriesOnTheSameLine)
        {
            //ExStart
            //ExFor:FieldIndex.RunSubentriesOnSameLine
            //ExSummary:Shows how to work with subentries in an INDEX field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an INDEX field which will display the page locations of XE fields in the document body
            FieldIndex index = (FieldIndex)builder.InsertField(FieldType.FieldIndex, true);

            // Normally, every XE field that's a subheading of any level is displayed on a unique line entry
            // in the INDEX field's table of contents
            // We can reduce the length of our INDEX table by putting all subheading entries along with their page locations on one line
            index.RunSubentriesOnSameLine = doRunSubentriesOnTheSameLine;
            index.PageNumberSeparator = ", see page ";
            index.Heading = "A";

            if (doRunSubentriesOnTheSameLine)
                Assert.AreEqual(" INDEX  \\r \\e \", see page \" \\h A", index.GetFieldCode());
            else
                Assert.AreEqual(" INDEX  \\e \", see page \" \\h A", index.GetFieldCode());

            // An XE field's "Text" attribute is the same thing as the "Heading" that will appear in the INDEX field's table of contents
            // This attribute can also contain one or multiple subheadings, separated by a colon (:),
            // which will be grouped under their parent headings/subheadings in the INDEX field
            // If index.RunSubentriesOnSameLine is false, "Heading 1" will take up one line as a heading,
            // followed by a two-line indented list of "Subheading 1" and "Subheading 2" with their respective page numbers
            // Otherwise, the two subheadings and their page numbers will be on the same line as their heading
            builder.InsertBreak(BreakType.PageBreak);
            FieldXE indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Heading 1:Subheading 1";

            Assert.AreEqual(" XE  \"Heading 1:Subheading 1\"", indexEntry.GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "Heading 1:Subheading 2";
            
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INDEX.XE.Subheading.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.INDEX.XE.Subheading.docx");
            index = (FieldIndex)doc.Range.Fields[0];

            if (doRunSubentriesOnTheSameLine)
            {
                TestUtil.VerifyField(FieldType.FieldIndex, " INDEX  \\r \\e \", see page \" \\h A",
                    "H\r" +
                    "Heading 1: Subheading 1, see page 2; Subheading 2, see page 3\r", index);
                Assert.True(index.RunSubentriesOnSameLine);
            }
            else
            {
                TestUtil.VerifyField(FieldType.FieldIndex, " INDEX  \\e \", see page \" \\h A",
                    "H\r" +
                    "Heading 1\r" +
                    "Subheading 1, see page 2\r" +
                    "Subheading 2, see page 3\r", index);
                Assert.False(index.RunSubentriesOnSameLine);
            }

            indexEntry = (FieldXE)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  \"Heading 1:Subheading 1\"", string.Empty, indexEntry);
            Assert.AreEqual("Heading 1:Subheading 1", indexEntry.Text);

            indexEntry = (FieldXE)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  \"Heading 1:Subheading 2\"", string.Empty, indexEntry);
            Assert.AreEqual("Heading 1:Subheading 2", indexEntry.Text);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Ignore("WORDSNET-17524")]
        public void FieldIndexYomi(bool doSortEntriesUsingYomi)
        {
            //ExStart
            //ExFor:FieldIndex.UseYomi
            //ExFor:FieldXE.Yomi
            //ExSummary:Shows how to sort INDEX field entries phonetically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an INDEX field which will display the page locations of XE fields in the document body
            FieldIndex index = (FieldIndex)builder.InsertField(FieldType.FieldIndex, true);

            // Set the INDEX table to sort entries phonetically using Hiragana
            index.UseYomi = doSortEntriesUsingYomi;

            if (doSortEntriesUsingYomi)
                Assert.AreEqual(" INDEX  \\y", index.GetFieldCode());
            else
                Assert.AreEqual(" INDEX ", index.GetFieldCode());

            // Insert 4 XE fields, which would show up as entries in the INDEX field's table of contents,
            // sorted in lexicographic order on their "Text" attribute
            builder.InsertBreak(BreakType.PageBreak);
            FieldXE indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "愛子";

            // The "Text" attribute may contain a word's spelling in Kanji, whose pronunciation may be ambiguous,
            // while a "Yomi" version of the word will be spelled exactly how it is pronounced using Hiragana
            // If our INDEX field is set to use Yomi, then we can sort phonetically using the "Yomi" attribute values instead of the "Text" attribute
            indexEntry.Yomi = "あ";

            Assert.AreEqual(" XE  愛子 \\y あ", indexEntry.GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "明美";
            indexEntry.Yomi = "あ";

            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "恵美";
            indexEntry.Yomi = "え";

            builder.InsertBreak(BreakType.PageBreak);
            indexEntry = (FieldXE)builder.InsertField(FieldType.FieldIndexEntry, true);
            indexEntry.Text = "愛美";
            indexEntry.Yomi = "え";

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.INDEX.XE.Yomi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.INDEX.XE.Yomi.docx");
            index = (FieldIndex)doc.Range.Fields[0];

            if (doSortEntriesUsingYomi)
            {
                Assert.True(index.UseYomi);
                Assert.AreEqual(" INDEX  \\y", index.GetFieldCode());
                Assert.AreEqual("愛子, 2\r" +
                                "明美, 3\r" +
                                "恵美, 4\r" +
                                "愛美, 5\r", index.Result);
            }
            else
            {
                Assert.False(index.UseYomi);
                Assert.AreEqual(" INDEX ", index.GetFieldCode());
                Assert.AreEqual("恵美, 4\r" +
                                "愛子, 2\r" +
                                "愛美, 5\r" +
                                "明美, 3\r", index.Result);
            }

            indexEntry = (FieldXE)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  愛子 \\y あ", string.Empty, indexEntry);
            Assert.AreEqual("愛子", indexEntry.Text);
            Assert.AreEqual("あ", indexEntry.Yomi);

            indexEntry = (FieldXE)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  明美 \\y あ", string.Empty, indexEntry);
            Assert.AreEqual("明美", indexEntry.Text);
            Assert.AreEqual("あ", indexEntry.Yomi);

            indexEntry = (FieldXE)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  恵美 \\y え", string.Empty, indexEntry);
            Assert.AreEqual("恵美", indexEntry.Text);
            Assert.AreEqual("え", indexEntry.Yomi);

            indexEntry = (FieldXE)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldIndexEntry, " XE  愛美 \\y え", string.Empty, indexEntry);
            Assert.AreEqual("愛美", indexEntry.Text);
            Assert.AreEqual("え", indexEntry.Yomi);
        }

        [Test]
        public void FieldBarcode()
        {
            //ExStart
            //ExFor:FieldBarcode
            //ExFor:FieldBarcode.FacingIdentificationMark
            //ExFor:FieldBarcode.IsBookmark
            //ExFor:FieldBarcode.IsUSPostalAddress
            //ExFor:FieldBarcode.PostalAddress
            //ExSummary:Shows how to insert a BARCODE field and set its properties. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a bookmark with a US postal code in it
            builder.StartBookmark("BarcodeBookmark");
            builder.Writeln("96801");
            builder.EndBookmark("BarcodeBookmark");

            builder.Writeln();

            // Reference a US postal code directly
            FieldBarcode field = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);
            field.FacingIdentificationMark = "C";
            field.PostalAddress = "96801";
            field.IsUSPostalAddress = true;

            Assert.AreEqual(" BARCODE  96801 \\f C \\u", field.GetFieldCode());

            builder.Writeln();

            // Reference a US postal code from a bookmark
            field = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);
            field.PostalAddress = "BarcodeBookmark";
            field.IsBookmark = true;

            Assert.AreEqual(" BARCODE  BarcodeBookmark \\b", field.GetFieldCode());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.BARCODE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.BARCODE.docx");

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true).Count);

            field = (FieldBarcode)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldBarcode, " BARCODE  96801 \\f C \\u", string.Empty, field);
            Assert.AreEqual("C", field.FacingIdentificationMark);
            Assert.AreEqual("96801", field.PostalAddress);
            Assert.True(field.IsUSPostalAddress);

            field = (FieldBarcode)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldBarcode, " BARCODE  BarcodeBookmark \\b", string.Empty, field);
            Assert.AreEqual("BarcodeBookmark", field.PostalAddress);
            Assert.True(field.IsBookmark);
        }

        [Test]
        public void FieldDisplayBarcode()
        {
            //ExStart
            //ExFor:FieldDisplayBarcode
            //ExFor:FieldDisplayBarcode.AddStartStopChar
            //ExFor:FieldDisplayBarcode.BackgroundColor
            //ExFor:FieldDisplayBarcode.BarcodeType
            //ExFor:FieldDisplayBarcode.BarcodeValue
            //ExFor:FieldDisplayBarcode.CaseCodeStyle
            //ExFor:FieldDisplayBarcode.DisplayText
            //ExFor:FieldDisplayBarcode.ErrorCorrectionLevel
            //ExFor:FieldDisplayBarcode.FixCheckDigit
            //ExFor:FieldDisplayBarcode.ForegroundColor
            //ExFor:FieldDisplayBarcode.PosCodeStyle
            //ExFor:FieldDisplayBarcode.ScalingFactor
            //ExFor:FieldDisplayBarcode.SymbolHeight
            //ExFor:FieldDisplayBarcode.SymbolRotation
            //ExSummary:Shows how to insert a DISPLAYBARCODE field and set its properties. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Insert a QR code
            field.BarcodeType = "QR";
            field.BarcodeValue = "ABC123";
            field.BackgroundColor = "0xF8BD69";
            field.ForegroundColor = "0xB5413B";
            field.ErrorCorrectionLevel = "3";
            field.ScalingFactor = "250";
            field.SymbolHeight = "1000";
            field.SymbolRotation = "0";

            Assert.AreEqual(" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", field.GetFieldCode());
            builder.Writeln();

            // Insert an EAN13 barcode
            field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            field.BarcodeType = "EAN13";
            field.BarcodeValue = "501234567890";
            field.DisplayText = true;
            field.PosCodeStyle = "CASE";
            field.FixCheckDigit = true;

            Assert.AreEqual(" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", field.GetFieldCode());
            builder.Writeln();

            // Insert a CODE39 barcode
            field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            field.BarcodeType = "CODE39";
            field.BarcodeValue = "12345ABCDE";
            field.AddStartStopChar = true;

            Assert.AreEqual(" DISPLAYBARCODE  12345ABCDE CODE39 \\d", field.GetFieldCode());
            builder.Writeln();

            // Insert an ITF14 barcode
            field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            field.BarcodeType = "ITF14";
            field.BarcodeValue = "09312345678907";
            field.CaseCodeStyle = "STD";

            Assert.AreEqual(" DISPLAYBARCODE  09312345678907 ITF14 \\c STD", field.GetFieldCode());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.DISPLAYBARCODE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.DISPLAYBARCODE.docx");

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true).Count);

            field = (FieldDisplayBarcode)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, " DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", string.Empty, field);
            Assert.AreEqual("QR", field.BarcodeType);
            Assert.AreEqual("ABC123", field.BarcodeValue);
            Assert.AreEqual("0xF8BD69", field.BackgroundColor);
            Assert.AreEqual("0xB5413B", field.ForegroundColor);
            Assert.AreEqual("3", field.ErrorCorrectionLevel);
            Assert.AreEqual("250", field.ScalingFactor);
            Assert.AreEqual("1000", field.SymbolHeight);
            Assert.AreEqual("0", field.SymbolRotation);

            field = (FieldDisplayBarcode)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, " DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", string.Empty, field);
            Assert.AreEqual("EAN13", field.BarcodeType);
            Assert.AreEqual("501234567890", field.BarcodeValue);
            Assert.True(field.DisplayText);
            Assert.AreEqual("CASE", field.PosCodeStyle);
            Assert.True(field.FixCheckDigit);

            field = (FieldDisplayBarcode)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, " DISPLAYBARCODE  12345ABCDE CODE39 \\d", string.Empty, field);
            Assert.AreEqual("CODE39", field.BarcodeType);
            Assert.AreEqual("12345ABCDE", field.BarcodeValue);
            Assert.True(field.AddStartStopChar);

            field = (FieldDisplayBarcode)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, " DISPLAYBARCODE  09312345678907 ITF14 \\c STD", string.Empty, field);
            Assert.AreEqual("ITF14", field.BarcodeType);
            Assert.AreEqual("09312345678907", field.BarcodeValue);
            Assert.AreEqual("STD", field.CaseCodeStyle);
        }


        [Test]
        public void FieldMergeBarcode_QR()
        {
            //ExStart
            //ExFor:FieldDisplayBarcode
            //ExFor:FieldMergeBarcode
            //ExFor:FieldMergeBarcode.BackgroundColor
            //ExFor:FieldMergeBarcode.BarcodeType
            //ExFor:FieldMergeBarcode.BarcodeValue
            //ExFor:FieldMergeBarcode.ErrorCorrectionLevel
            //ExFor:FieldMergeBarcode.ForegroundColor
            //ExFor:FieldMergeBarcode.ScalingFactor
            //ExFor:FieldMergeBarcode.SymbolHeight
            //ExFor:FieldMergeBarcode.SymbolRotation
            //ExSummary:Shows how to perform a mail merge on QR barcodes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEBARCODE field,
            // which functions like a MERGEFIELD by creating a barcode from the merged data source's values
            // This field will convert all rows in a merge data source's "MyQRCode" column into QR barcodes
            FieldMergeBarcode field = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
            field.BarcodeType = "QR";
            field.BarcodeValue = "MyQRCode";

            // Edit its appearance such as colors and scale
            field.BackgroundColor = "0xF8BD69";
            field.ForegroundColor = "0xB5413B";
            field.ErrorCorrectionLevel = "3";
            field.ScalingFactor = "250";
            field.SymbolHeight = "1000";
            field.SymbolRotation = "0";

            Assert.AreEqual(FieldType.FieldMergeBarcode, field.Type);
            Assert.AreEqual(" MERGEBARCODE  MyQRCode QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0",
                field.GetFieldCode());
            builder.Writeln();

            // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue
            // When we execute the mail merge,
            // a barcode of a type we specified in the MERGEBARCODE field will be created with each row's value
            DataTable table = new DataTable("Barcodes");
            table.Columns.Add("MyQRCode");
            table.Rows.Add(new[] { "ABC123" });
            table.Rows.Add(new[] { "DEF456" });

            doc.MailMerge.Execute(table);

            // Every row in the "MyQRCode" column has created a DISPLAYBARCODE field, which shows a barcode with the merged value
            Assert.AreEqual(FieldType.FieldDisplayBarcode, doc.Range.Fields[0].Type);
            Assert.AreEqual("DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", 
                doc.Range.Fields[0].GetFieldCode());
            Assert.AreEqual(FieldType.FieldDisplayBarcode, doc.Range.Fields[1].Type);
            Assert.AreEqual("DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B",
                doc.Range.Fields[1].GetFieldCode());

            doc.Save(ArtifactsDir + "Field.MERGEBARCODE.QR.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.MERGEBARCODE.QR.docx");

            Assert.AreEqual(0, doc.Range.Fields.Count(f => f.Type == FieldType.FieldMergeBarcode));

            FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, 
                "DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", string.Empty, barcode);
            Assert.AreEqual("ABC123", barcode.BarcodeValue);
            Assert.AreEqual("QR", barcode.BarcodeType);

            barcode = (FieldDisplayBarcode)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, 
                "DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", string.Empty, barcode);
            Assert.AreEqual("DEF456", barcode.BarcodeValue);
            Assert.AreEqual("QR", barcode.BarcodeType);
        }

        [Test]
        public void FieldMergeBarcode_EAN13()
        {
            //ExStart
            //ExFor:FieldMergeBarcode
            //ExFor:FieldMergeBarcode.BarcodeType
            //ExFor:FieldMergeBarcode.BarcodeValue
            //ExFor:FieldMergeBarcode.DisplayText
            //ExFor:FieldMergeBarcode.FixCheckDigit
            //ExFor:FieldMergeBarcode.PosCodeStyle
            //ExSummary:Shows how to perform a mail merge on EAN13 barcodes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEBARCODE field,
            // which functions like a MERGEFIELD by creating a barcode from the merged data source's values
            // This field will convert all rows in a merge data source's "MyEAN13Barcode" column into EAN13 barcodes
            FieldMergeBarcode field = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
            field.BarcodeType = "EAN13";
            field.BarcodeValue = "MyEAN13Barcode";

            // Edit its appearance to display barcode data under the lines
            field.DisplayText = true;
            field.PosCodeStyle = "CASE";
            field.FixCheckDigit = true;

            Assert.AreEqual(FieldType.FieldMergeBarcode, field.Type);
            Assert.AreEqual(" MERGEBARCODE  MyEAN13Barcode EAN13 \\t \\p CASE \\x", field.GetFieldCode());
            builder.Writeln();

            // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue
            // When we execute the mail merge,
            // a barcode of a type we specified in the MERGEBARCODE field will be created with each row's value
            DataTable table = new DataTable("Barcodes");
            table.Columns.Add("MyEAN13Barcode");
            table.Rows.Add(new[] { "501234567890" });
            table.Rows.Add(new[] { "123456789012" });

            doc.MailMerge.Execute(table);

            // Every row in the "MyEAN13Barcode" column has created a DISPLAYBARCODE field,
            // which shows a barcode with the merged value
            Assert.AreEqual(FieldType.FieldDisplayBarcode, doc.Range.Fields[0].Type);
            Assert.AreEqual("DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x",
                doc.Range.Fields[0].GetFieldCode());
            Assert.AreEqual(FieldType.FieldDisplayBarcode, doc.Range.Fields[1].Type);
            Assert.AreEqual("DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x",
                doc.Range.Fields[1].GetFieldCode());

            doc.Save(ArtifactsDir + "Field.MERGEBARCODE.EAN13.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.MERGEBARCODE.EAN13.docx");

            Assert.AreEqual(0, doc.Range.Fields.Count(f => f.Type == FieldType.FieldMergeBarcode));

            FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x", string.Empty, barcode);
            Assert.AreEqual("501234567890", barcode.BarcodeValue);
            Assert.AreEqual("EAN13", barcode.BarcodeType);

            barcode = (FieldDisplayBarcode)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x", string.Empty, barcode);
            Assert.AreEqual("123456789012", barcode.BarcodeValue);
            Assert.AreEqual("EAN13", barcode.BarcodeType);
        }

        [Test]
        public void FieldMergeBarcode_CODE39()
        {
            //ExStart
            //ExFor:FieldMergeBarcode
            //ExFor:FieldMergeBarcode.AddStartStopChar
            //ExFor:FieldMergeBarcode.BarcodeType
            //ExSummary:Shows how to perform a mail merge on CODE39 barcodes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEBARCODE field,
            // which functions like a MERGEFIELD by creating a barcode from the merged data source's values
            // This field will convert all rows in a merge data source's "MyCODE39Barcode" column into CODE39 barcodes
            FieldMergeBarcode field = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
            field.BarcodeType = "CODE39";
            field.BarcodeValue = "MyCODE39Barcode";

            // Edit its appearance to display start/stop characters
            field.AddStartStopChar = true;

            Assert.AreEqual(FieldType.FieldMergeBarcode, field.Type);
            Assert.AreEqual(" MERGEBARCODE  MyCODE39Barcode CODE39 \\d", field.GetFieldCode());
            builder.Writeln();

            // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue
            // When we execute the mail merge,
            // a barcode of a type we specified in the MERGEBARCODE field will be created with each row's value
            DataTable table = new DataTable("Barcodes");
            table.Columns.Add("MyCODE39Barcode");
            table.Rows.Add(new[] { "12345ABCDE" });
            table.Rows.Add(new[] { "67890FGHIJ" });

            doc.MailMerge.Execute(table);

            // Every row in the "MyCODE39Barcode" column has created a DISPLAYBARCODE field,
            // which shows a barcode with the merged value
            Assert.AreEqual(FieldType.FieldDisplayBarcode, doc.Range.Fields[0].Type);
            Assert.AreEqual("DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d",
                doc.Range.Fields[0].GetFieldCode());
            Assert.AreEqual(FieldType.FieldDisplayBarcode, doc.Range.Fields[1].Type);
            Assert.AreEqual("DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d",
                doc.Range.Fields[1].GetFieldCode());

            doc.Save(ArtifactsDir + "Field.MERGEBARCODE.CODE39.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.MERGEBARCODE.CODE39.docx");

            Assert.AreEqual(0, doc.Range.Fields.Count(f => f.Type == FieldType.FieldMergeBarcode));

            FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d", string.Empty, barcode);
            Assert.AreEqual("12345ABCDE", barcode.BarcodeValue);
            Assert.AreEqual("CODE39", barcode.BarcodeType);

            barcode = (FieldDisplayBarcode)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d", string.Empty, barcode);
            Assert.AreEqual("67890FGHIJ", barcode.BarcodeValue);
            Assert.AreEqual("CODE39", barcode.BarcodeType);
        }

        [Test]
        public void FieldMergeBarcode_ITF14()
        {
            //ExStart
            //ExFor:FieldMergeBarcode
            //ExFor:FieldMergeBarcode.BarcodeType
            //ExFor:FieldMergeBarcode.CaseCodeStyle
            //ExSummary:Shows how to perform a mail merge on ITF14 barcodes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEBARCODE field,
            // which functions like a MERGEFIELD by creating a barcode from the merged data source's values
            // This field will convert all rows in a merge data source's "MyITF14Barcode" column into ITF14 barcodes
            FieldMergeBarcode field = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
            field.BarcodeType = "ITF14";
            field.BarcodeValue = "MyITF14Barcode";
            field.CaseCodeStyle = "STD";

            Assert.AreEqual(FieldType.FieldMergeBarcode, field.Type);
            Assert.AreEqual(" MERGEBARCODE  MyITF14Barcode ITF14 \\c STD", field.GetFieldCode());

            // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue
            // When we execute the mail merge,
            // a barcode of a type we specified in the MERGEBARCODE field will be created with each row's value
            DataTable table = new DataTable("Barcodes");
            table.Columns.Add("MyITF14Barcode");
            table.Rows.Add(new[] { "09312345678907" });
            table.Rows.Add(new[] { "1234567891234" });

            doc.MailMerge.Execute(table);

            // Every row in the "MyITF14Barcode" column has created a DISPLAYBARCODE field,
            // which shows a barcode with the merged value
            Assert.AreEqual(FieldType.FieldDisplayBarcode, doc.Range.Fields[0].Type);
            Assert.AreEqual("DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD",
                doc.Range.Fields[0].GetFieldCode());
            Assert.AreEqual(FieldType.FieldDisplayBarcode, doc.Range.Fields[1].Type);
            Assert.AreEqual("DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD",
                doc.Range.Fields[1].GetFieldCode());

            doc.Save(ArtifactsDir + "Field.MERGEBARCODE.ITF14.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.MERGEBARCODE.ITF14.docx");

            Assert.AreEqual(0, doc.Range.Fields.Count(f => f.Type == FieldType.FieldMergeBarcode));

            FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD", string.Empty, barcode);
            Assert.AreEqual("09312345678907", barcode.BarcodeValue);
            Assert.AreEqual("ITF14", barcode.BarcodeType);

            barcode = (FieldDisplayBarcode)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD", string.Empty, barcode);
            Assert.AreEqual("1234567891234", barcode.BarcodeValue);
            Assert.AreEqual("ITF14", barcode.BarcodeType);
        }

        //ExStart
        //ExFor:FieldLink
        //ExFor:FieldLink.AutoUpdate
        //ExFor:FieldLink.FormatUpdateType
        //ExFor:FieldLink.InsertAsBitmap
        //ExFor:FieldLink.InsertAsHtml
        //ExFor:FieldLink.InsertAsPicture
        //ExFor:FieldLink.InsertAsRtf
        //ExFor:FieldLink.InsertAsText
        //ExFor:FieldLink.InsertAsUnicode
        //ExFor:FieldLink.IsLinked
        //ExFor:FieldLink.ProgId
        //ExFor:FieldLink.SourceFullName
        //ExFor:FieldLink.SourceItem
        //ExFor:FieldDde
        //ExFor:FieldDde.AutoUpdate
        //ExFor:FieldDde.InsertAsBitmap
        //ExFor:FieldDde.InsertAsHtml
        //ExFor:FieldDde.InsertAsPicture
        //ExFor:FieldDde.InsertAsRtf
        //ExFor:FieldDde.InsertAsText
        //ExFor:FieldDde.InsertAsUnicode
        //ExFor:FieldDde.IsLinked
        //ExFor:FieldDde.ProgId
        //ExFor:FieldDde.SourceFullName
        //ExFor:FieldDde.SourceItem
        //ExFor:FieldDdeAuto
        //ExFor:FieldDdeAuto.InsertAsBitmap
        //ExFor:FieldDdeAuto.InsertAsHtml
        //ExFor:FieldDdeAuto.InsertAsPicture
        //ExFor:FieldDdeAuto.InsertAsRtf
        //ExFor:FieldDdeAuto.InsertAsText
        //ExFor:FieldDdeAuto.InsertAsUnicode
        //ExFor:FieldDdeAuto.IsLinked
        //ExFor:FieldDdeAuto.ProgId
        //ExFor:FieldDdeAuto.SourceFullName
        //ExFor:FieldDdeAuto.SourceItem
        //ExSummary:Shows how to insert linked objects as LINK, DDE and DDEAUTO fields and present them within the document in different ways.
        [TestCase(InsertLinkedObjectAs.Text)] //ExSkip
        [TestCase(InsertLinkedObjectAs.Unicode)] //ExSkip
        [TestCase(InsertLinkedObjectAs.Html)] //ExSkip
        [TestCase(InsertLinkedObjectAs.Rtf)] //ExSkip
        [Ignore("WORDSNET-16226")] //ExSkip
        public void FieldLinkedObjectsAsText(InsertLinkedObjectAs insertLinkedObjectAs)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert fields containing text from another document and present them as text (see InsertLinkedObjectAs enum)
            builder.Writeln("FieldLink:\n");
            InsertFieldLink(builder, insertLinkedObjectAs, "Word.Document.8", MyDir + "Document.docx", null, true);

            builder.Writeln("FieldDde:\n");
            InsertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", MyDir + "Spreadsheet.xlsx",
                "Sheet1!R1C1", true, true);

            builder.Writeln("FieldDdeAuto:\n");
            InsertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", MyDir + "Spreadsheet.xlsx",
                "Sheet1!R1C1", true);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.LINK.DDE.DDEAUTO.docx");
        }

        [TestCase(InsertLinkedObjectAs.Picture)] //ExSkip
        [TestCase(InsertLinkedObjectAs.Bitmap)] //ExSkip
        [Ignore("WORDSNET-16226")] //ExSkip
        public void FieldLinkedObjectsAsImage(InsertLinkedObjectAs insertLinkedObjectAs)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert one cell from a spreadsheet as an image (see InsertLinkedObjectAs enum)
            builder.Writeln("FieldLink:\n");
            InsertFieldLink(builder, insertLinkedObjectAs, "Excel.Sheet", MyDir + "MySpreadsheet.xlsx",
                "Sheet1!R2C2", true);

            builder.Writeln("FieldDde:\n");
            InsertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", MyDir + "Spreadsheet.xlsx",
                "Sheet1!R1C1", true, true);

            builder.Writeln("FieldDdeAuto:\n");
            InsertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", MyDir + "Spreadsheet.xlsx",
                "Sheet1!R1C1", true);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.LINK.DDE.DDEAUTO.AsImage.docx");
        }

        /// <summary>
        /// Use a document builder to insert a LINK field and set its properties according to parameters.
        /// </summary>
        private static void InsertFieldLink(DocumentBuilder builder, InsertLinkedObjectAs insertLinkedObjectAs,
            string progId, string sourceFullName, string sourceItem, bool shouldAutoUpdate)
        {
            FieldLink field = (FieldLink)builder.InsertField(FieldType.FieldLink, true);

            switch (insertLinkedObjectAs)
            {
                case InsertLinkedObjectAs.Text:
                    field.InsertAsText = true;
                    break;
                case InsertLinkedObjectAs.Unicode:
                    field.InsertAsUnicode = true;
                    break;
                case InsertLinkedObjectAs.Html:
                    field.InsertAsHtml = true;
                    break;
                case InsertLinkedObjectAs.Rtf:
                    field.InsertAsRtf = true;
                    break;
                case InsertLinkedObjectAs.Picture:
                    field.InsertAsPicture = true;
                    break;
                case InsertLinkedObjectAs.Bitmap:
                    field.InsertAsBitmap = true;
                    break;
            }

            field.AutoUpdate = shouldAutoUpdate;
            field.ProgId = progId;
            field.SourceFullName = sourceFullName;
            field.SourceItem = sourceItem;

            builder.Writeln("\n");
        }

        /// <summary>
        /// Use a document builder to insert a DDE field and set its properties according to parameters.
        /// </summary>
        private static void InsertFieldDde(DocumentBuilder builder, InsertLinkedObjectAs insertLinkedObjectAs, string progId,
            string sourceFullName, string sourceItem, bool isLinked, bool shouldAutoUpdate)
        {
            FieldDde field = (FieldDde)builder.InsertField(FieldType.FieldDDE, true);

            switch (insertLinkedObjectAs)
            {
                case InsertLinkedObjectAs.Text:
                    field.InsertAsText = true;
                    break;
                case InsertLinkedObjectAs.Unicode:
                    field.InsertAsUnicode = true;
                    break;
                case InsertLinkedObjectAs.Html:
                    field.InsertAsHtml = true;
                    break;
                case InsertLinkedObjectAs.Rtf:
                    field.InsertAsRtf = true;
                    break;
                case InsertLinkedObjectAs.Picture:
                    field.InsertAsPicture = true;
                    break;
                case InsertLinkedObjectAs.Bitmap:
                    field.InsertAsBitmap = true;
                    break;
            }

            field.AutoUpdate = shouldAutoUpdate;
            field.ProgId = progId;
            field.SourceFullName = sourceFullName;
            field.SourceItem = sourceItem;
            field.IsLinked = isLinked;

            builder.Writeln("\n");
        }

        /// <summary>
        /// Use a document builder to insert a DDEAUTO field and set its properties according to parameters.
        /// </summary>
        private static void InsertFieldDdeAuto(DocumentBuilder builder, InsertLinkedObjectAs insertLinkedObjectAs,
            string progId, string sourceFullName, string sourceItem, bool isLinked)
        {
            FieldDdeAuto field = (FieldDdeAuto)builder.InsertField(FieldType.FieldDDEAuto, true);

            switch (insertLinkedObjectAs)
            {
                case InsertLinkedObjectAs.Text:
                    field.InsertAsText = true;
                    break;
                case InsertLinkedObjectAs.Unicode:
                    field.InsertAsUnicode = true;
                    break;
                case InsertLinkedObjectAs.Html:
                    field.InsertAsHtml = true;
                    break;
                case InsertLinkedObjectAs.Rtf:
                    field.InsertAsRtf = true;
                    break;
                case InsertLinkedObjectAs.Picture:
                    field.InsertAsPicture = true;
                    break;
                case InsertLinkedObjectAs.Bitmap:
                    field.InsertAsBitmap = true;
                    break;
            }

            field.ProgId = progId;
            field.SourceFullName = sourceFullName;
            field.SourceItem = sourceItem;
            field.IsLinked = isLinked;
        }

        public enum InsertLinkedObjectAs
        {
            // LinkedObjectAsText
            Text,
            Unicode,
            Html,
            Rtf,
            // LinkedObjectAsImage
            Picture,
            Bitmap
        }
        //ExEnd

        [Test]
        public void FieldUserAddress()
        {
            //ExStart
            //ExFor:FieldUserAddress
            //ExFor:FieldUserAddress.UserAddress
            //ExSummary:Shows how to use the USERADDRESS field.
            Document doc = new Document();

            // Create a user information object and set it as the data source for our field
            UserInformation userInformation = new UserInformation();
            userInformation.Address = "123 Main Street";
            doc.FieldOptions.CurrentUser = userInformation;

            // Display the current user's address with a USERADDRESS field
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldUserAddress fieldUserAddress = (FieldUserAddress)builder.InsertField(FieldType.FieldUserAddress, true);
            Assert.AreEqual(userInformation.Address, fieldUserAddress.Result);

            Assert.AreEqual(" USERADDRESS ", fieldUserAddress.GetFieldCode());
            Assert.AreEqual("123 Main Street", fieldUserAddress.Result);

            // We can set this attribute to get our field to display a different value
            fieldUserAddress.UserAddress = "456 North Road";
            fieldUserAddress.Update();

            Assert.AreEqual(" USERADDRESS  \"456 North Road\"", fieldUserAddress.GetFieldCode());
            Assert.AreEqual("456 North Road", fieldUserAddress.Result);

            // This does not change the value in the user information object
            Assert.AreEqual("123 Main Street", doc.FieldOptions.CurrentUser.Address);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.USERADDRESS.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.USERADDRESS.docx");

            fieldUserAddress = (FieldUserAddress)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldUserAddress, " USERADDRESS  \"456 North Road\"", "456 North Road", fieldUserAddress);
            Assert.AreEqual("456 North Road", fieldUserAddress.UserAddress);
        }

        [Test]
        public void FieldUserInitials()
        {
            //ExStart
            //ExFor:FieldUserInitials
            //ExFor:FieldUserInitials.UserInitials
            //ExSummary:Shows how to use the USERINITIALS field.
            Document doc = new Document();

            // Create a user information object and set it as the data source for our field
            UserInformation userInformation = new UserInformation();
            userInformation.Initials = "J. D.";
            doc.FieldOptions.CurrentUser = userInformation;

            // Display the current user's Initials with a USERINITIALS field
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldUserInitials fieldUserInitials = (FieldUserInitials)builder.InsertField(FieldType.FieldUserInitials, true);
            Assert.AreEqual(userInformation.Initials, fieldUserInitials.Result);

            Assert.AreEqual(" USERINITIALS ", fieldUserInitials.GetFieldCode());
            Assert.AreEqual("J. D.", fieldUserInitials.Result);

            // We can set this attribute to get our field to display a different value
            fieldUserInitials.UserInitials = "J. C.";
            fieldUserInitials.Update();

            Assert.AreEqual(" USERINITIALS  \"J. C.\"", fieldUserInitials.GetFieldCode());
            Assert.AreEqual("J. C.", fieldUserInitials.Result);

            // This does not change the value in the user information object
            Assert.AreEqual("J. D.", doc.FieldOptions.CurrentUser.Initials);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.USERINITIALS.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.USERINITIALS.docx");

            fieldUserInitials = (FieldUserInitials)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldUserInitials, " USERINITIALS  \"J. C.\"", "J. C.", fieldUserInitials);
            Assert.AreEqual("J. C.", fieldUserInitials.UserInitials);
        }

        [Test]
        public void FieldUserName()
        {
            //ExStart
            //ExFor:FieldUserName
            //ExFor:FieldUserName.UserName
            //ExSummary:Shows how to use the USERNAME field.
            Document doc = new Document();

            // Create a user information object and set it as the data source for our field
            UserInformation userInformation = new UserInformation();
            userInformation.Name = "John Doe";
            doc.FieldOptions.CurrentUser = userInformation;

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Display the current user's Name with a USERNAME field
            FieldUserName fieldUserName = (FieldUserName)builder.InsertField(FieldType.FieldUserName, true);
            Assert.AreEqual(userInformation.Name, fieldUserName.Result);

            Assert.AreEqual(" USERNAME ", fieldUserName.GetFieldCode());
            Assert.AreEqual("John Doe", fieldUserName.Result);

            // We can set this attribute to get our field to display a different value
            fieldUserName.UserName = "Jane Doe";
            fieldUserName.Update();

            Assert.AreEqual(" USERNAME  \"Jane Doe\"", fieldUserName.GetFieldCode());
            Assert.AreEqual("Jane Doe", fieldUserName.Result);

            // This does not change the value in the user information object
            Assert.AreEqual("John Doe", doc.FieldOptions.CurrentUser.Name);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.USERNAME.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.USERNAME.docx");

            fieldUserName = (FieldUserName)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldUserName, " USERNAME  \"Jane Doe\"", "Jane Doe", fieldUserName);
            Assert.AreEqual("Jane Doe", fieldUserName.UserName);
        }

        [Test]
        [Ignore("WORDSNET-17657")]
        public void FieldStyleRefParagraphNumbers()
        {
            //ExStart
            //ExFor:FieldStyleRef
            //ExFor:FieldStyleRef.InsertParagraphNumber
            //ExFor:FieldStyleRef.InsertParagraphNumberInFullContext
            //ExFor:FieldStyleRef.InsertParagraphNumberInRelativeContext
            //ExFor:FieldStyleRef.InsertRelativePosition
            //ExFor:FieldStyleRef.SearchFromBottom
            //ExFor:FieldStyleRef.StyleName
            //ExFor:FieldStyleRef.SuppressNonDelimiters
            //ExSummary:Shows how to use STYLEREF fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list based on one of the Microsoft Word list templates
            Aspose.Words.Lists.List list = doc.Lists.Add(Aspose.Words.Lists.ListTemplate.NumberDefault);

            // This generated list will look like "1.a )"
            // The space before the bracket is a non-delimiter character that can be suppressed
            list.ListLevels[0].NumberFormat = "\x0000.";
            list.ListLevels[1].NumberFormat = "\x0001 )";

            // Add text and apply paragraph styles that will be referenced by STYLEREF fields
            builder.ListFormat.List = list;
            builder.ListFormat.ListIndent();
            builder.ParagraphFormat.Style = doc.Styles["List Paragraph"];
            builder.Writeln("Item 1");
            builder.ParagraphFormat.Style = doc.Styles["Quote"];
            builder.Writeln("Item 2");
            builder.ParagraphFormat.Style = doc.Styles["List Paragraph"];
            builder.Writeln("Item 3");
            builder.ListFormat.RemoveNumbers();
            builder.ParagraphFormat.Style = doc.Styles["Normal"];

            // Place a STYLEREF field in the header and have it display the first "List Paragraph"-styled text in the document
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            FieldStyleRef field = (FieldStyleRef)builder.InsertField(FieldType.FieldStyleRef, true);
            field.StyleName = "List Paragraph";

            // Place a STYLEREF field in the footer and have it display the last text
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            field = (FieldStyleRef)builder.InsertField(FieldType.FieldStyleRef, true);
            field.StyleName = "List Paragraph";
            field.SearchFromBottom = true;

            builder.MoveToDocumentEnd();

            // We can also use STYLEREF fields to reference the list numbers of lists
            builder.Write("\nParagraph number: ");
            field = (FieldStyleRef)builder.InsertField(FieldType.FieldStyleRef, true);
            field.StyleName = "Quote";
            field.InsertParagraphNumber = true;

            builder.Write("\nParagraph number, relative context: ");
            field = (FieldStyleRef)builder.InsertField(FieldType.FieldStyleRef, true);
            field.StyleName = "Quote";
            field.InsertParagraphNumberInRelativeContext = true;

            builder.Write("\nParagraph number, full context: ");
            field = (FieldStyleRef)builder.InsertField(FieldType.FieldStyleRef, true);
            field.StyleName = "Quote";
            field.InsertParagraphNumberInFullContext = true;

            builder.Write("\nParagraph number, full context, non-delimiter chars suppressed: ");
            field = (FieldStyleRef)builder.InsertField(FieldType.FieldStyleRef, true);
            field.StyleName = "Quote";
            field.InsertParagraphNumberInFullContext = true;
            field.SuppressNonDelimiters = true;

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.STYLEREF.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.STYLEREF.docx");

            field = (FieldStyleRef)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldStyleRef, " STYLEREF  \"List Paragraph\"", "Item 1", field);
            Assert.AreEqual("List Paragraph", field.StyleName);

            field = (FieldStyleRef)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldStyleRef, " STYLEREF  \"List Paragraph\" \\l", "Item 3", field);
            Assert.AreEqual("List Paragraph", field.StyleName);
            Assert.True(field.SearchFromBottom);

            field = (FieldStyleRef)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldStyleRef, " STYLEREF  Quote \\n", "b )", field);
            Assert.AreEqual("Quote", field.StyleName);
            Assert.True(field.InsertParagraphNumber);

            field = (FieldStyleRef)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldStyleRef, " STYLEREF  Quote \\r", "b )", field);
            Assert.AreEqual("Quote", field.StyleName);
            Assert.True(field.InsertParagraphNumberInRelativeContext);

            field = (FieldStyleRef)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldStyleRef, " STYLEREF  Quote \\w", "1.b )", field);
            Assert.AreEqual("Quote", field.StyleName);
            Assert.True(field.InsertParagraphNumberInFullContext);

            field = (FieldStyleRef)doc.Range.Fields[5];

            TestUtil.VerifyField(FieldType.FieldStyleRef, " STYLEREF  Quote \\w \\t", "1.b)", field);
            Assert.AreEqual("Quote", field.StyleName);
            Assert.True(field.InsertParagraphNumberInFullContext);
            Assert.True(field.SuppressNonDelimiters);
        }

#if NET462 || NETCOREAPP2_1 || JAVA
        [Test]
        public void FieldDate()
        {
            //ExStart
            //ExFor:FieldDate
            //ExFor:FieldDate.UseLunarCalendar
            //ExFor:FieldDate.UseSakaEraCalendar
            //ExFor:FieldDate.UseUmAlQuraCalendar
            //ExFor:FieldDate.UseLastFormat
            //ExSummary:Shows how to insert DATE fields with different kinds of calendars.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // One way of putting dates into our documents is inserting DATE fields with document builder
            FieldDate field = (FieldDate)builder.InsertField(FieldType.FieldDate, true);

            // Set the field's date to the current date of the Islamic Lunar Calendar
            field.UseLunarCalendar = true;
            Assert.AreEqual(" DATE  \\h", field.GetFieldCode());
            builder.Writeln();

            // Insert a date field with the current date of the Umm al-Qura calendar
            field = (FieldDate)builder.InsertField(FieldType.FieldDate, true);
            field.UseUmAlQuraCalendar = true;
            Assert.AreEqual(" DATE  \\u", field.GetFieldCode());
            builder.Writeln();

            // Insert a date field with the current date of the Indian national calendar
            field = (FieldDate)builder.InsertField(FieldType.FieldDate, true);
            field.UseSakaEraCalendar = true;
            Assert.AreEqual(" DATE  \\s", field.GetFieldCode());
            builder.Writeln();

            // Insert a date field with the current date of the calendar used in the (Insert > Date and Time) dialog box
            field = (FieldDate)builder.InsertField(FieldType.FieldDate, true);
            field.UseLastFormat = true;
            Assert.AreEqual(" DATE  \\l", field.GetFieldCode());
            builder.Writeln();

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.DATE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.DATE.docx");

            field = (FieldDate)doc.Range.Fields[0];

            Assert.AreEqual(FieldType.FieldDate, field.Type);
            Assert.True(field.UseLunarCalendar);
            Assert.AreEqual(" DATE  \\h", field.GetFieldCode());
            Assert.IsTrue(Regex.Match(doc.Range.Fields[0].Result, @"\d{1,2}[/]\d{1,2}[/]\d{4}").Success);

            field = (FieldDate)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldDate, " DATE  \\u", DateTime.Now.ToShortDateString(), field);
            Assert.True(field.UseUmAlQuraCalendar);

            field = (FieldDate)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldDate, " DATE  \\s", DateTime.Now.ToShortDateString(), field);
            Assert.True(field.UseSakaEraCalendar);

            field = (FieldDate)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldDate, " DATE  \\l", DateTime.Now.ToShortDateString(), field);
            Assert.True(field.UseLastFormat);
        }
#endif

        [Test]
        [Ignore("WORDSNET-17669")]
        public void FieldCreateDate()
        {
            //ExStart
            //ExFor:FieldCreateDate
            //ExFor:FieldCreateDate.UseLunarCalendar
            //ExFor:FieldCreateDate.UseSakaEraCalendar
            //ExFor:FieldCreateDate.UseUmAlQuraCalendar
            //ExSummary:Shows how to insert CREATEDATE fields to display document creation dates.
            // Open an existing document and move a document builder to the end
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln(" Date this document was created:");

            // Insert a CREATEDATE field and display, using the Lunar Calendar, the date the document was created
            builder.Write("According to the Lunar Calendar - ");
            FieldCreateDate field = (FieldCreateDate)builder.InsertField(FieldType.FieldCreateDate, true);
            field.UseLunarCalendar = true;

            Assert.AreEqual(" CREATEDATE  \\h", field.GetFieldCode());

            // Display the date using the Umm al-Qura Calendar
            builder.Write("\nAccording to the Umm al-Qura Calendar - ");
            field = (FieldCreateDate)builder.InsertField(FieldType.FieldCreateDate, true);
            field.UseUmAlQuraCalendar = true;

            Assert.AreEqual(" CREATEDATE  \\u", field.GetFieldCode());

            // Display the date using the Indian National Calendar
            builder.Write("\nAccording to the Indian National Calendar - ");
            field = (FieldCreateDate)builder.InsertField(FieldType.FieldCreateDate, true);
            field.UseSakaEraCalendar = true;

            Assert.AreEqual(" CREATEDATE  \\s", field.GetFieldCode());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.CREATEDATE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.CREATEDATE.docx");

            Assert.AreEqual(new DateTime(2017, 12, 5, 9, 56, 0), doc.BuiltInDocumentProperties.CreatedTime);

            DateTime expectedDate = doc.BuiltInDocumentProperties.CreatedTime.AddHours(TimeZoneInfo.Local.GetUtcOffset(DateTime.UtcNow).Hours);
            field = (FieldCreateDate)doc.Range.Fields[0];
            Calendar umAlQuraCalendar = new UmAlQuraCalendar();

            TestUtil.VerifyField(FieldType.FieldCreateDate, " CREATEDATE  \\h",
                $"{umAlQuraCalendar.GetMonth(expectedDate)}/{umAlQuraCalendar.GetDayOfMonth(expectedDate)}/{umAlQuraCalendar.GetYear(expectedDate)} " +
                expectedDate.AddHours(1).ToString("hh:mm:ss tt"), field);
            Assert.AreEqual(FieldType.FieldCreateDate, field.Type);
            Assert.True(field.UseLunarCalendar);
            
            field = (FieldCreateDate)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldCreateDate, " CREATEDATE  \\u",
                $"{umAlQuraCalendar.GetMonth(expectedDate)}/{umAlQuraCalendar.GetDayOfMonth(expectedDate)}/{umAlQuraCalendar.GetYear(expectedDate)} " +
                expectedDate.AddHours(1).ToString("hh:mm:ss tt"), field);
            Assert.AreEqual(FieldType.FieldCreateDate, field.Type);
            Assert.True(field.UseUmAlQuraCalendar);
        }

        [Test]
        [Ignore("WORDSNET-17669")]
        public void FieldSaveDate()
        {
            //ExStart
            //ExFor:FieldSaveDate
            //ExFor:FieldSaveDate.UseLunarCalendar
            //ExFor:FieldSaveDate.UseSakaEraCalendar
            //ExFor:FieldSaveDate.UseUmAlQuraCalendar
            //ExSummary:Shows how to insert SAVEDATE fields the date and time a document was last saved.
            // Open an existing document and move a document builder to the end
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln(" Date this document was last saved:");

            // Insert a SAVEDATE field and display, using the Lunar Calendar, the date the document was last saved
            builder.Write("According to the Lunar Calendar - ");
            FieldSaveDate field = (FieldSaveDate)builder.InsertField(FieldType.FieldSaveDate, true);
            field.UseLunarCalendar = true;

            Assert.AreEqual(" SAVEDATE  \\h", field.GetFieldCode());
            
            // Display the date using the Umm al-Qura Calendar
            builder.Write("\nAccording to the Umm al-Qura calendar - ");
            field = (FieldSaveDate)builder.InsertField(FieldType.FieldSaveDate, true);
            field.UseUmAlQuraCalendar = true;

            Assert.AreEqual(" SAVEDATE  \\u", field.GetFieldCode());

            // Display the date using the Indian National Calendar
            builder.Write("\nAccording to the Indian National calendar - ");
            field = (FieldSaveDate)builder.InsertField(FieldType.FieldSaveDate, true);
            field.UseSakaEraCalendar = true;

            Assert.AreEqual(" SAVEDATE  \\s", field.GetFieldCode());
            
            // While the date/time of the most recent save operation is tracked automatically by Microsoft Word,
            // we will need to update the value manually if we wish to do the same thing when calling the Save() method
            doc.BuiltInDocumentProperties.LastSavedTime = DateTime.Now;

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.SAVEDATE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SAVEDATE.docx");

            Console.WriteLine(doc.BuiltInDocumentProperties.LastSavedTime);

            field = (FieldSaveDate)doc.Range.Fields[0];

            Assert.AreEqual(FieldType.FieldSaveDate, field.Type);
            Assert.True(field.UseLunarCalendar);
            Assert.AreEqual(" SAVEDATE  \\h", field.GetFieldCode());

            Assert.True(Regex.Match(field.Result, "\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M").Success);

            field = (FieldSaveDate)doc.Range.Fields[1];

            Assert.AreEqual(FieldType.FieldSaveDate, field.Type);
            Assert.True(field.UseUmAlQuraCalendar);
            Assert.AreEqual(" SAVEDATE  \\u", field.GetFieldCode());
            Assert.True(Regex.Match(field.Result, "\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M").Success);
        }

        [Test]
        public void FieldBuilder()
        {
            //ExStart
            //ExFor:FieldBuilder
            //ExFor:FieldBuilder.AddArgument(Int32)
            //ExFor:FieldBuilder.AddArgument(FieldArgumentBuilder)
            //ExFor:FieldBuilder.AddArgument(String)
            //ExFor:FieldBuilder.AddArgument(Double)
            //ExFor:FieldBuilder.AddArgument(FieldBuilder)
            //ExFor:FieldBuilder.AddSwitch(String)
            //ExFor:FieldBuilder.AddSwitch(String, Double)
            //ExFor:FieldBuilder.AddSwitch(String, Int32)
            //ExFor:FieldBuilder.AddSwitch(String, String)
            //ExFor:FieldBuilder.BuildAndInsert(Paragraph)
            //ExFor:FieldArgumentBuilder
            //ExFor:FieldArgumentBuilder.AddField(FieldBuilder)
            //ExFor:FieldArgumentBuilder.AddText(String)
            //ExFor:FieldArgumentBuilder.AddNode(Inline)
            //ExSummary:Shows how to insert fields using a field builder.
            Document doc = new Document();

            // Use a field builder to add a SYMBOL field which displays the "F with hook" symbol
            FieldBuilder builder = new FieldBuilder(FieldType.FieldSymbol);
            builder.AddArgument(402);
            builder.AddSwitch("\\f", "Arial");
            builder.AddSwitch("\\s", 25);
            builder.AddSwitch("\\u");
            Field field = builder.BuildAndInsert(doc.FirstSection.Body.FirstParagraph);

            Assert.AreEqual(" SYMBOL 402 \\f Arial \\s 25 \\u ", field.GetFieldCode());

            // Use a field builder to create a formula field that will be used by another field builder
            FieldBuilder innerFormulaBuilder = new FieldBuilder(FieldType.FieldFormula);
            innerFormulaBuilder.AddArgument(100);
            innerFormulaBuilder.AddArgument("+");
            innerFormulaBuilder.AddArgument(74);

            // Add a field builder as an argument to another field builder
            // The result of our formula field will be used as an ANSI value representing the "enclosed R" symbol,
            // to be displayed by this SYMBOL field
            builder = new FieldBuilder(FieldType.FieldSymbol);
            builder.AddArgument(innerFormulaBuilder);
            field = builder.BuildAndInsert(doc.FirstSection.Body.AppendParagraph(""));

            Assert.AreEqual(" SYMBOL \u0013 = 100 + 74 \u0014\u0015 ", field.GetFieldCode());

            // Now we will use our builder to construct a more complex field with nested fields
            // For our IF field, we will first create two formula fields to serve as expressions
            // Their results will be tested for equality to decide what value an IF field displays
            FieldBuilder leftExpression = new FieldBuilder(FieldType.FieldFormula);
            leftExpression.AddArgument(2);
            leftExpression.AddArgument("+");
            leftExpression.AddArgument(3);

            FieldBuilder rightExpression = new FieldBuilder(FieldType.FieldFormula);
            rightExpression.AddArgument(2.5);
            rightExpression.AddArgument("*");
            rightExpression.AddArgument(5.2);

            // Next, we will create two field arguments using field argument builders
            // These will serve as the two possible outputs of our IF field and they will also use our two expressions
            FieldArgumentBuilder trueOutput = new FieldArgumentBuilder();
            trueOutput.AddText("True, both expressions amount to ");
            trueOutput.AddField(leftExpression);

            FieldArgumentBuilder falseOutput = new FieldArgumentBuilder();
            falseOutput.AddNode(new Run(doc, "False, "));
            falseOutput.AddField(leftExpression);
            falseOutput.AddNode(new Run(doc, " does not equal "));
            falseOutput.AddField(rightExpression);

            // Finally, we will use a field builder to create an IF field which takes two field builders as expressions,
            // and two field argument builders as the two potential outputs
            builder = new FieldBuilder(FieldType.FieldIf);
            builder.AddArgument(leftExpression);
            builder.AddArgument("=");
            builder.AddArgument(rightExpression);
            builder.AddArgument(trueOutput);
            builder.AddArgument(falseOutput);

            builder.BuildAndInsert(doc.FirstSection.Body.AppendParagraph(""));

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.SYMBOL.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SYMBOL.docx");

            FieldSymbol fieldSymbol = (FieldSymbol)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldSymbol, " SYMBOL 402 \\f Arial \\s 25 \\u ", string.Empty, fieldSymbol);
            Assert.AreEqual("ƒ", fieldSymbol.DisplayResult);

            fieldSymbol = (FieldSymbol)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldSymbol, " SYMBOL \u0013 = 100 + 74 \u0014174\u0015 ", string.Empty, fieldSymbol);
            Assert.AreEqual("®", fieldSymbol.DisplayResult);

            TestUtil.VerifyField(FieldType.FieldFormula, " = 100 + 74 ", "174", doc.Range.Fields[2]);

            TestUtil.VerifyField(FieldType.FieldIf,
                " IF \u0013 = 2 + 3 \u00145\u0015 = \u0013 = 2.5 * 5.2 \u001413\u0015 " +
                "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
                "\"False, \u0013 = 2 + 3 \u00145\u0015 does not equal \u0013 = 2.5 * 5.2 \u001413\u0015\" ",
                "False, 5 does not equal 13", doc.Range.Fields[3]);

            Assert.Throws<AssertionException>(() => TestUtil.FieldsAreNested(doc.Range.Fields[2], doc.Range.Fields[3]));

            TestUtil.VerifyField(FieldType.FieldFormula, " = 2 + 3 ", "5", doc.Range.Fields[4]);
            TestUtil.FieldsAreNested(doc.Range.Fields[4], doc.Range.Fields[3]);

            TestUtil.VerifyField(FieldType.FieldFormula, " = 2.5 * 5.2 ", "13", doc.Range.Fields[5]);
            TestUtil.FieldsAreNested(doc.Range.Fields[5], doc.Range.Fields[3]);

            TestUtil.VerifyField(FieldType.FieldFormula, " = 2 + 3 ", string.Empty, doc.Range.Fields[6]);
            TestUtil.FieldsAreNested(doc.Range.Fields[6], doc.Range.Fields[3]);

            TestUtil.VerifyField(FieldType.FieldFormula, " = 2 + 3 ", "5", doc.Range.Fields[7]);
            TestUtil.FieldsAreNested(doc.Range.Fields[7], doc.Range.Fields[3]);

            TestUtil.VerifyField(FieldType.FieldFormula, " = 2.5 * 5.2 ", "13", doc.Range.Fields[8]);
            TestUtil.FieldsAreNested(doc.Range.Fields[8], doc.Range.Fields[3]);
        }
        
        [Test]
        public void FieldAuthor()
        {
            //ExStart
            //ExFor:FieldAuthor
            //ExFor:FieldAuthor.AuthorName  
            //ExFor:FieldOptions.DefaultDocumentAuthor
            //ExSummary:Shows how to display a document creator's name with an AUTHOR field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // If we open an existing document, the document's author's full name will be displayed by the field
            // If we create a document programmatically, we need to set this attribute to the author's name, so our field has something to display
            doc.FieldOptions.DefaultDocumentAuthor = "Joe Bloggs";

            builder.Write("This document was created by ");
            FieldAuthor field = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);
            field.Update();

            Assert.AreEqual(" AUTHOR ", field.GetFieldCode());
            Assert.AreEqual("Joe Bloggs", field.Result);
            
            // If this property has a value, it will supersede the one we set above 
            doc.BuiltInDocumentProperties.Author = "John Doe";      
            field.Update();

            Assert.AreEqual(" AUTHOR ", field.GetFieldCode());
            Assert.AreEqual("John Doe", field.Result);
            
            // Our field can also override the document's built in author name like this
            field.AuthorName = "Jane Doe";
            field.Update();

            Assert.AreEqual(" AUTHOR  \"Jane Doe\"", field.GetFieldCode());
            Assert.AreEqual("Jane Doe", field.Result);

            // The author name in the built-in properties was changed by the field, but the default document author stays the same
            Assert.AreEqual("Jane Doe", doc.BuiltInDocumentProperties.Author);
            Assert.AreEqual("Joe Bloggs", doc.FieldOptions.DefaultDocumentAuthor);

            doc.Save(ArtifactsDir + "Field.AUTHOR.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.AUTHOR.docx");

            Assert.Null(doc.FieldOptions.DefaultDocumentAuthor);
            Assert.AreEqual("Jane Doe", doc.BuiltInDocumentProperties.Author);

            field = (FieldAuthor)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldAuthor, " AUTHOR  \"Jane Doe\"", "Jane Doe", field);
            Assert.AreEqual("Jane Doe", field.AuthorName);
        }

        [Test]
        public void FieldDocVariable()
        {
            //ExStart
            //ExFor:FieldDocProperty
            //ExFor:FieldDocVariable
            //ExFor:FieldDocVariable.VariableName
            //ExSummary:Shows how to use fields to display document properties and variables.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the value of a document property
            doc.BuiltInDocumentProperties.Category = "My category";

            // Display the value of that property with a DOCPROPERTY field
            FieldDocProperty fieldDocProperty = (FieldDocProperty)builder.InsertField(" DOCPROPERTY Category ");
            fieldDocProperty.Update();

            Assert.AreEqual(" DOCPROPERTY Category ", fieldDocProperty.GetFieldCode());
            Assert.AreEqual("My category", fieldDocProperty.Result);

            builder.Writeln();

            // While the set of a document's properties is fixed, we can add, name, and define our own values in the variables collection
            Assert.That(doc.Variables, Is.Empty);
            doc.Variables.Add("My variable", "My variable's value");

            // We can access a variable using its name and display it with a DOCVARIABLE field
            FieldDocVariable fieldDocVariable = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
            fieldDocVariable.VariableName = "My Variable";
            fieldDocVariable.Update();

            Assert.AreEqual(" DOCVARIABLE  \"My Variable\"", fieldDocVariable.GetFieldCode());
            Assert.AreEqual("My variable's value", fieldDocVariable.Result);

            doc.Save(ArtifactsDir + "Field.DOCPROPERTY.DOCVARIABLE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.DOCPROPERTY.DOCVARIABLE.docx");

            Assert.AreEqual("My category", doc.BuiltInDocumentProperties.Category);

            fieldDocProperty = (FieldDocProperty)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldDocProperty, " DOCPROPERTY Category ", "My category", fieldDocProperty);

            fieldDocVariable = (FieldDocVariable)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldDocVariable, " DOCVARIABLE  \"My Variable\"", "My variable's value", fieldDocVariable);
            Assert.AreEqual("My Variable", fieldDocVariable.VariableName);
        }

        [Test]
        public void FieldSubject()
        {
            //ExStart
            //ExFor:FieldSubject
            //ExFor:FieldSubject.Text
            //ExSummary:Shows how to use the SUBJECT field.
            Document doc = new Document();

            // Set a value for the document's subject property
            doc.BuiltInDocumentProperties.Subject = "My subject";

            // We can display this value with a SUBJECT field
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldSubject field = (FieldSubject)builder.InsertField(FieldType.FieldSubject, true);
            field.Update();

            Assert.AreEqual(" SUBJECT ", field.GetFieldCode());
            Assert.AreEqual("My subject", field.Result);

            // We can also set the field's Text attribute to override the current value of the Subject property
            field.Text = "My new subject";
            field.Update();

            Assert.AreEqual(" SUBJECT  \"My new subject\"", field.GetFieldCode());
            Assert.AreEqual("My new subject", field.Result);

            // As well as displaying a new value in our field, we also changed the value of the document property
            Assert.AreEqual("My new subject", doc.BuiltInDocumentProperties.Subject);

            doc.Save(ArtifactsDir + "Field.SUBJECT.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SUBJECT.docx");

            Assert.AreEqual("My new subject", doc.BuiltInDocumentProperties.Subject);

            field = (FieldSubject)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldSubject, " SUBJECT  \"My new subject\"", "My new subject", field);
            Assert.AreEqual("My new subject", field.Text);
        }

        [Test]
        public void FieldComments()
        {
            //ExStart
            //ExFor:FieldComments
            //ExFor:FieldComments.Text
            //ExSummary:Shows how to use the COMMENTS field to display a document's comments.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // This property is where the COMMENTS field will source its content from
            doc.BuiltInDocumentProperties.Comments = "My comment.";

            // Insert a COMMENTS field with a document builder
            FieldComments field = (FieldComments)builder.InsertField(FieldType.FieldComments, true);
            field.Update();

            Assert.AreEqual(" COMMENTS ", field.GetFieldCode());
            Assert.AreEqual("My comment.", field.Result);

            // We can override the comment from the document's built in properties and display any text we put here instead
            field.Text = "My overriding comment.";
            field.Update();

            Assert.AreEqual(" COMMENTS  \"My overriding comment.\"", field.GetFieldCode());
            Assert.AreEqual("My overriding comment.", field.Result);

            doc.Save(ArtifactsDir + "Field.COMMENTS.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.COMMENTS.docx");

            Assert.AreEqual("My overriding comment.", doc.BuiltInDocumentProperties.Comments);

            field = (FieldComments)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldComments, " COMMENTS  \"My overriding comment.\"", "My overriding comment.", field);
            Assert.AreEqual("My overriding comment.", field.Text);
        }
        
        [Test]
        public void FieldFileSize()
        {
            //ExStart
            //ExFor:FieldFileSize
            //ExFor:FieldFileSize.IsInKilobytes
            //ExFor:FieldFileSize.IsInMegabytes            
            //ExSummary:Shows how to display the file size of a document with a FILESIZE field.
            // Open a document and verify its file size
            Document doc = new Document(MyDir + "Document.docx");

            Assert.AreEqual(10590, doc.BuiltInDocumentProperties.Bytes);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.InsertParagraph();

            // By default, file size is displayed in bytes
            FieldFileSize field = (FieldFileSize)builder.InsertField(FieldType.FieldFileSize, true);
            field.Update();

            Assert.AreEqual(" FILESIZE ", field.GetFieldCode());
            Assert.AreEqual("10590", field.Result);

            // Set the field to display size in kilobytes
            builder.InsertParagraph();
            field = (FieldFileSize)builder.InsertField(FieldType.FieldFileSize, true);
            field.IsInKilobytes = true;
            field.Update();

            Assert.AreEqual(" FILESIZE  \\k", field.GetFieldCode());
            Assert.AreEqual("11", field.Result);

            // Set the field to display size in megabytes
            builder.InsertParagraph();
            field = (FieldFileSize)builder.InsertField(FieldType.FieldFileSize, true);
            field.IsInMegabytes = true;
            field.Update();

            Assert.AreEqual(" FILESIZE  \\m", field.GetFieldCode());
            Assert.AreEqual("0", field.Result);

            // To update the values of these fields while editing in Microsoft Word,
            // the changes must first be saved, then the fields need to be manually updated
            doc.Save(ArtifactsDir + "Field.FILESIZE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.FILESIZE.docx");

            field = (FieldFileSize)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldFileSize, " FILESIZE ", "10590", field);

            // These fields will need to be updated to produce an accurate result
            doc.UpdateFields();

            field = (FieldFileSize)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldFileSize, " FILESIZE  \\k", "9", field);
            Assert.True(field.IsInKilobytes);

            field = (FieldFileSize)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldFileSize, " FILESIZE  \\m", "0", field);
            Assert.True(field.IsInMegabytes);
        }

        [Test]
        public void FieldGoToButton()
        {
            //ExStart
            //ExFor:FieldGoToButton
            //ExFor:FieldGoToButton.DisplayText
            //ExFor:FieldGoToButton.Location
            //ExSummary:Shows to insert a GOTOBUTTON field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a GOTOBUTTON which will take us to a bookmark referenced by "MyBookmark"
            FieldGoToButton field = (FieldGoToButton)builder.InsertField(FieldType.FieldGoToButton, true);
            field.DisplayText = "My Button";
            field.Location = "MyBookmark";

            Assert.AreEqual(" GOTOBUTTON  MyBookmark My Button", field.GetFieldCode());

            // Add an arrival destination for our button
            builder.InsertBreak(BreakType.PageBreak);
            builder.StartBookmark(field.Location);
            builder.Writeln("Bookmark text contents.");
            builder.EndBookmark(field.Location);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.GOTOBUTTON.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.GOTOBUTTON.docx");
            field = (FieldGoToButton)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldGoToButton, " GOTOBUTTON  MyBookmark My Button", string.Empty, field);
            Assert.AreEqual("My Button", field.DisplayText);
            Assert.AreEqual("MyBookmark", field.Location);
        }
        
        [Test]
        //ExStart
        //ExFor:FieldFillIn
        //ExFor:FieldFillIn.DefaultResponse
        //ExFor:FieldFillIn.PromptOnceOnMailMerge
        //ExFor:FieldFillIn.PromptText
        //ExSummary:Shows how to use the FILLIN field to prompt the user for a response.
        public void FieldFillIn()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a FILLIN field with a document builder
            FieldFillIn field = (FieldFillIn)builder.InsertField(FieldType.FieldFillIn, true);
            field.PromptText = "Please enter a response:";
            field.DefaultResponse = "A default response.";

            // Set this to prompt the user for a response when a mail merge is performed
            field.PromptOnceOnMailMerge = true;

            Assert.AreEqual(" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", field.GetFieldCode());

            // Perform a simple mail merge
            FieldMergeField mergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            mergeField.FieldName = "MergeField";
            
            doc.FieldOptions.UserPromptRespondent = new PromptRespondent();
            doc.MailMerge.Execute(new [] { "MergeField" }, new object[] { "" });
            
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.FILLIN.docx");
            TestFieldFillIn(new Document(ArtifactsDir + "Field.FILLIN.docx")); //ExSKip
        }

        /// <summary>
        /// IFieldUserPromptRespondent implementation that appends a line to the default response of an FILLIN field during a mail merge.
        /// </summary>
        private class PromptRespondent : IFieldUserPromptRespondent
        {
            public string Respond(string promptText, string defaultResponse)
            {
                return "Response modified by PromptRespondent. " + defaultResponse;
            }
        }
        //ExEnd

        private void TestFieldFillIn(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual(1, doc.Range.Fields.Count);

            FieldFillIn field = (FieldFillIn)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldFillIn, " FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", 
                "Response modified by PromptRespondent. A default response.", field);
            Assert.AreEqual("Please enter a response:", field.PromptText);
            Assert.AreEqual("A default response.", field.DefaultResponse);
            Assert.True(field.PromptOnceOnMailMerge);
        }

        [Test]
        public void FieldInfo()
        {
            //ExStart
            //ExFor:FieldInfo
            //ExFor:FieldInfo.InfoType
            //ExFor:FieldInfo.NewValue
            //ExSummary:Shows how to work with INFO fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the value of a document property
            doc.BuiltInDocumentProperties.Comments = "My comment";

            // We can access a property using its name and display it with an INFO field
            // In this case, it will be the Comments property
            FieldInfo field = (FieldInfo)builder.InsertField(FieldType.FieldInfo, true);
            field.InfoType = "Comments";
            field.Update();

            Assert.AreEqual(" INFO  Comments", field.GetFieldCode());
            Assert.AreEqual("My comment", field.Result);

            builder.Writeln();

            // We can override the value of a document property by setting an INFO field's optional new value
            field = (FieldInfo)builder.InsertField(FieldType.FieldInfo, true);
            field.InfoType = "Comments";
            field.NewValue = "New comment";
            field.Update();

            // Our field's new value has been applied to the corresponding property
            Assert.AreEqual(" INFO  Comments \"New comment\"", field.GetFieldCode());
            Assert.AreEqual("New comment", field.Result);
            Assert.AreEqual("New comment", doc.BuiltInDocumentProperties.Comments);

            doc.Save(ArtifactsDir + "Field.INFO.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.INFO.docx");

            Assert.AreEqual("New comment", doc.BuiltInDocumentProperties.Comments);
            
            field = (FieldInfo)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldInfo, " INFO  Comments", "My comment", field);
            Assert.AreEqual("Comments", field.InfoType);

            field = (FieldInfo)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldInfo, " INFO  Comments \"New comment\"", "New comment", field);
            Assert.AreEqual("Comments", field.InfoType);
            Assert.AreEqual("New comment", field.NewValue);
        }

        [Test]
        public void FieldMacroButton()
        {
            //ExStart
            //ExFor:Document.HasMacros
            //ExFor:FieldMacroButton
            //ExFor:FieldMacroButton.DisplayText
            //ExFor:FieldMacroButton.MacroName
            //ExSummary:Shows how to use MACROBUTTON fields that enable us to run macros by clicking.
            // Open a document that contains macros
            Document doc = new Document(MyDir + "Macro.docm");
            DocumentBuilder builder = new DocumentBuilder(doc);

            Assert.IsTrue(doc.HasMacros);

            // Insert a MACROBUTTON field and reference by name a macro that exists within the input document
            FieldMacroButton field = (FieldMacroButton)builder.InsertField(FieldType.FieldMacroButton, true);
            field.MacroName = "MyMacro";
            field.DisplayText = "Double click to run macro: " + field.MacroName;

            Assert.AreEqual(" MACROBUTTON  MyMacro Double click to run macro: MyMacro", field.GetFieldCode());

            // Reference "ViewZoom200", a macro that was shipped with Microsoft Word, found under "Word commands"
            // If our document has a macro of the same name as one from another source, the field will select ours to run
            builder.InsertParagraph();
            field = (FieldMacroButton)builder.InsertField(FieldType.FieldMacroButton, true);
            field.MacroName = "ViewZoom200";
            field.DisplayText = "Run " + field.MacroName;

            Assert.AreEqual(" MACROBUTTON  ViewZoom200 Run ViewZoom200", field.GetFieldCode());

            // Save the document as a macro-enabled document type
            doc.Save(ArtifactsDir + "Field.MACROBUTTON.docm");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.MACROBUTTON.docm");

            field = (FieldMacroButton)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldMacroButton, " MACROBUTTON  MyMacro Double click to run macro: MyMacro", string.Empty, field);
            Assert.AreEqual("MyMacro", field.MacroName);
            Assert.AreEqual("Double click to run macro: MyMacro", field.DisplayText);

            field = (FieldMacroButton)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldMacroButton, " MACROBUTTON  ViewZoom200 Run ViewZoom200", string.Empty, field);
            Assert.AreEqual("ViewZoom200", field.MacroName);
            Assert.AreEqual("Run ViewZoom200", field.DisplayText);
        }

        [Test]
        public void FieldKeywords()
        {
            //ExStart
            //ExFor:FieldKeywords
            //ExFor:FieldKeywords.Text
            //ExSummary:Shows to insert a KEYWORDS field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some keywords, also referred to as "tags" in File Explorer
            doc.BuiltInDocumentProperties.Keywords = "Keyword1, Keyword2";

            // Add a KEYWORDS field which will display our keywords
            FieldKeywords field = (FieldKeywords)builder.InsertField(FieldType.FieldKeyword, true);
            field.Update();

            Assert.AreEqual(" KEYWORDS ", field.GetFieldCode());
            Assert.AreEqual("Keyword1, Keyword2", field.Result);

            // We can set the Text property of our field to display a different value to the one within the document's properties
            field.Text = "OverridingKeyword";
            field.Update();

            Assert.AreEqual(" KEYWORDS  OverridingKeyword", field.GetFieldCode());
            Assert.AreEqual("OverridingKeyword", field.Result);

            // Setting a KEYWORDS field's Text property also updates the document's keywords to our new value
            Assert.AreEqual("OverridingKeyword", doc.BuiltInDocumentProperties.Keywords);

            doc.Save(ArtifactsDir + "Field.KEYWORDS.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.KEYWORDS.docx");

            Assert.AreEqual("OverridingKeyword", doc.BuiltInDocumentProperties.Keywords);

            field = (FieldKeywords)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldKeyword, " KEYWORDS  OverridingKeyword", "OverridingKeyword", field);
            Assert.AreEqual("OverridingKeyword", field.Text);
        }

        [Test]
        public void FieldNum()
        {
            //ExStart
            //ExFor:FieldPage
            //ExFor:FieldNumChars
            //ExFor:FieldNumPages
            //ExFor:FieldNumWords
            //ExSummary:Shows how to use NUMCHARS, NUMWORDS, NUMPAGES and PAGE fields to track the size of our documents.
            // Open a document to which we want to add character/word/page counts
            Document doc = new Document(MyDir + "Paragraphs.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the document builder to the footer, where we will store our fields
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Insert character and word counts
            FieldNumChars fieldNumChars = (FieldNumChars)builder.InsertField(FieldType.FieldNumChars, true);       
            builder.Writeln(" characters");
            FieldNumWords fieldNumWords = (FieldNumWords)builder.InsertField(FieldType.FieldNumWords, true);
            builder.Writeln(" words");

            // Insert a "Page x of y" page count
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Page ");
            FieldPage fieldPage = (FieldPage)builder.InsertField(FieldType.FieldPage, true);
            builder.Write(" of ");
            FieldNumPages fieldNumPages = (FieldNumPages)builder.InsertField(FieldType.FieldNumPages, true);

            Assert.AreEqual(" NUMCHARS ", fieldNumChars.GetFieldCode());
            Assert.AreEqual(" NUMWORDS ", fieldNumWords.GetFieldCode());
            Assert.AreEqual(" NUMPAGES ", fieldNumPages.GetFieldCode());
            Assert.AreEqual(" PAGE ", fieldPage.GetFieldCode());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");

            TestUtil.VerifyField(FieldType.FieldNumChars, " NUMCHARS ", "6009", doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldNumWords, " NUMWORDS ", "1054", doc.Range.Fields[1]);
            TestUtil.VerifyField(FieldType.FieldPage, " PAGE ", "6", doc.Range.Fields[2]);
            TestUtil.VerifyField(FieldType.FieldNumPages, " NUMPAGES ", "6", doc.Range.Fields[3]);
        }

        [Test]
        public void FieldPrint()
        {
            //ExStart
            //ExFor:FieldPrint
            //ExFor:FieldPrint.PostScriptGroup
            //ExFor:FieldPrint.PrinterInstructions
            //ExSummary:Shows to insert a PRINT field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("My paragraph");

            // The PRINT field can send instructions to the printer that we use to print our document
            FieldPrint field = (FieldPrint)builder.InsertField(FieldType.FieldPrint, true);

            // Set the area for the printer to perform instructions over
            // In this case, it will be the paragraph that contains our PRINT field
            field.PostScriptGroup = "para";

            // When our document is printed using a printer that supports PostScript,
            // this command will turn the entire area that we specified in field.PostScriptGroup white 
            field.PrinterInstructions = "erasepage";

            Assert.AreEqual(" PRINT  erasepage \\p para", field.GetFieldCode());
            
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.PRINT.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.PRINT.docx");

            field = (FieldPrint)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldPrint, " PRINT  erasepage \\p para", string.Empty, field);
            Assert.AreEqual("para", field.PostScriptGroup);
            Assert.AreEqual("erasepage", field.PrinterInstructions);
        }

        [Test]
        public void FieldPrintDate()
        {
            //ExStart
            //ExFor:FieldPrintDate
            //ExFor:FieldPrintDate.UseLunarCalendar
            //ExFor:FieldPrintDate.UseSakaEraCalendar
            //ExFor:FieldPrintDate.UseUmAlQuraCalendar
            //ExSummary:Shows read PRINTDATE fields.
            Document doc = new Document(MyDir + "Field sample - PRINTDATE.docx");
            
            // A PRINTDATE field will display "0/0/0000" by default
            // When a document is printed by a printer or printed as a PDF (but not exported as PDF),
            // these fields will display the date/time of that print operation
            FieldPrintDate field = (FieldPrintDate)doc.Range.Fields[0];

            Assert.AreEqual("3/25/2020 12:00:00 AM", field.Result);
            Assert.AreEqual(" PRINTDATE ", field.GetFieldCode());

            // These fields can also display the date using other various international calendars
            field = (FieldPrintDate)doc.Range.Fields[1];

            Assert.True(field.UseLunarCalendar);
            Assert.AreEqual("8/1/1441 12:00:00 AM", field.Result);
            Assert.AreEqual(" PRINTDATE  \\h", field.GetFieldCode());

            field = (FieldPrintDate)doc.Range.Fields[2];

            Assert.True(field.UseUmAlQuraCalendar);
            Assert.AreEqual("8/1/1441 12:00:00 AM", field.Result);
            Assert.AreEqual(" PRINTDATE  \\u", field.GetFieldCode());

            field = (FieldPrintDate)doc.Range.Fields[3];

            Assert.True(field.UseSakaEraCalendar);
            Assert.AreEqual("1/5/1942 12:00:00 AM", field.Result);
            Assert.AreEqual(" PRINTDATE  \\s", field.GetFieldCode());
            //ExEnd
        }

        [Test]
        public void FieldQuote()
        {
            //ExStart
            //ExFor:FieldQuote
            //ExFor:FieldQuote.Text
            //ExFor:Document.UpdateFields
            //ExSummary:Shows to use the QUOTE field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a QUOTE field, which will display content from the Text attribute
            FieldQuote field = (FieldQuote)builder.InsertField(FieldType.FieldQuote, true);
            field.Text = "\"Quoted text\"";

            Assert.AreEqual(" QUOTE  \"\\\"Quoted text\\\"\"", field.GetFieldCode());

            // Insert a QUOTE field with a nested DATE field
            // DATE fields normally update their value to the current date every time the document is opened
            // Nesting the DATE field inside the QUOTE field like this will freeze its value to the date when we created the document
            builder.Write("\nDocument creation date: ");
            field = (FieldQuote)builder.InsertField(FieldType.FieldQuote, true);
            builder.MoveTo(field.Separator);
            builder.InsertField(FieldType.FieldDate, true);

            Assert.AreEqual(" QUOTE \u0013 DATE \u0014" + DateTime.Now.Date.ToShortDateString() + "\u0015", field.GetFieldCode());

            // Some field types don't display the correct result until they are manually updated
            Assert.AreEqual(string.Empty, doc.Range.Fields[0].Result); 

            doc.UpdateFields();

            Assert.AreEqual("\"Quoted text\"", doc.Range.Fields[0].Result);

            doc.Save(ArtifactsDir + "Field.QUOTE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.QUOTE.docx");

            TestUtil.VerifyField(FieldType.FieldQuote, " QUOTE  \"\\\"Quoted text\\\"\"", "\"Quoted text\"", doc.Range.Fields[0]);

            TestUtil.VerifyField(FieldType.FieldQuote, " QUOTE \u0013 DATE \u0014" + DateTime.Now.Date.ToShortDateString() + "\u0015", 
                DateTime.Now.Date.ToShortDateString(), doc.Range.Fields[1]);

        }

        //ExStart
        //ExFor:FieldNext
        //ExFor:FieldNextIf
        //ExFor:FieldNextIf.ComparisonOperator
        //ExFor:FieldNextIf.LeftExpression
        //ExFor:FieldNextIf.RightExpression
        //ExSummary:Shows how to use NEXT/NEXTIF fields to merge more than one row into one page during a mail merge.
        [Test] //ExSkip
        public void FieldNext()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a data source for our mail merge with 3 rows,
            // This would normally amount to 3 pages in the output of a mail merge
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Courtesy Title");
            table.Columns.Add("First Name");
            table.Columns.Add("Last Name");
            table.Rows.Add("Mr.", "John", "Doe");
            table.Rows.Add("Mrs.", "Jane", "Cardholder");
            table.Rows.Add("Mr.", "Joe", "Bloggs");

            // Insert a set of merge fields
            InsertMergeFields(builder, "First row: ");

            // If we have multiple merge fields with the same FieldName,
            // they will receive data from the same row of the data source and will display the same value after the merge
            // A NEXT field tells the mail merge instantly to move down one row,
            // so any upcoming merge fields will have data deposited from the next row
            // Make sure not to skip with a NEXT/NEXTIF field while on the last row
            FieldNext fieldNext = (FieldNext)builder.InsertField(FieldType.FieldNext, true);

            Assert.AreEqual(" NEXT ", fieldNext.GetFieldCode());

            // These merge fields are the same as the ones as above but will take values from the second row
            InsertMergeFields(builder, "Second row: ");

            // A NEXTIF field has the same function as a NEXT field,
            // but it skips to the next row only if a condition expressed by the following 3 attributes is fulfilled
            FieldNextIf fieldNextIf = (FieldNextIf)builder.InsertField(FieldType.FieldNextIf, true);
            fieldNextIf.LeftExpression = "5";
            fieldNextIf.RightExpression = "2 + 3";
            fieldNextIf.ComparisonOperator = "=";

            // If the comparison asserted by the above field is correct,
            // the following 3 merge fields will take data from the third row
            // Otherwise, these fields will take data from row 2 again 
            InsertMergeFields(builder, "Third row: ");

            // Our data source has 3 rows and we skipped rows twice, so our output will have one page
            // with data from all 3 rows
            doc.MailMerge.Execute(table);

            Assert.AreEqual(" NEXTIF  5 = \"2 + 3\"", fieldNextIf.GetFieldCode());

            doc.Save(ArtifactsDir + "Field.NEXT.NEXTIF.docx");
            TestFieldNext(doc); //ExSKip
        }

        /// <summary>
        /// Uses a document builder to insert merge fields for a data table that has "Courtesy Title", "First Name" and "Last Name" columns.
        /// </summary>
        public void InsertMergeFields(DocumentBuilder builder, string firstFieldTextBefore)
        {
            InsertMergeField(builder, "Courtesy Title", firstFieldTextBefore, " ");
            InsertMergeField(builder, "First Name", null, " ");
            InsertMergeField(builder, "Last Name", null, null);
            builder.InsertParagraph();
        }

        /// <summary>
        /// Uses a document builder to insert a merge field.
        /// </summary>
        public void InsertMergeField(DocumentBuilder builder, string fieldName, string textBefore, string textAfter)
        {
            FieldMergeField field = (FieldMergeField) builder.InsertField(FieldType.FieldMergeField, true);
            field.FieldName = fieldName;
            field.TextBefore = textBefore;
            field.TextAfter = textAfter;
        }
        //ExEnd

        private void TestFieldNext(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual(0, doc.Range.Fields.Count);
            Assert.AreEqual("First row: Mr. John Doe\r" +
                            "Second row: Mrs. Jane Cardholder\r" +
                            "Third row: Mr. Joe Bloggs\r\f", doc.GetText());
        }

        //ExStart
        //ExFor:FieldNoteRef
        //ExFor:FieldNoteRef.BookmarkName
        //ExFor:FieldNoteRef.InsertHyperlink
        //ExFor:FieldNoteRef.InsertReferenceMark
        //ExFor:FieldNoteRef.InsertRelativePosition
        //ExSummary:Shows to insert NOTEREF fields and modify their appearance.
        [Test] //ExSkip
        [Ignore("WORDSNET-17845")] //ExSkip
        public void FieldNoteRef()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a bookmark with a footnote for the NOTEREF field to reference
            InsertBookmarkWithFootnote(builder, "MyBookmark1", "Contents of MyBookmark1", "Footnote from MyBookmark1");

            // This NOTEREF field will display just the number of the footnote inside the referenced bookmark
            // Setting the InsertHyperlink attribute lets us jump to the bookmark by Ctrl + clicking the field
            Assert.AreEqual(" NOTEREF  MyBookmark2 \\h",
                InsertFieldNoteRef(builder, "MyBookmark2", true, false, false, "Hyperlink to Bookmark2, with footnote number ").GetFieldCode());

            // When using the \p flag, after the footnote number the field also displays the position of the bookmark relative to the field
            // Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update
            Assert.AreEqual(" NOTEREF  MyBookmark1 \\h \\p",
                InsertFieldNoteRef(builder, "MyBookmark1", true, true, false, "Bookmark1, with footnote number ").GetFieldCode());

            // Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below"
            // The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text
            Assert.AreEqual(" NOTEREF  MyBookmark2 \\h \\p \\f",
                InsertFieldNoteRef(builder, "MyBookmark2", true, true, true, "Bookmark2, with footnote number ").GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);
            InsertBookmarkWithFootnote(builder, "MyBookmark2", "Contents of MyBookmark2", "Footnote from MyBookmark2");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.NOTEREF.docx");
            TestNoteRef(new Document(ArtifactsDir + "Field.NOTEREF.docx")); //ExSkip
        }

        /// <summary>
        /// Uses a document builder to insert a NOTEREF field and sets its attributes.
        /// </summary>
        private static FieldNoteRef InsertFieldNoteRef(DocumentBuilder builder, string bookmarkName, bool insertHyperlink, bool insertRelativePosition, bool insertReferenceMark, string textBefore)
        {
            builder.Write(textBefore);

            FieldNoteRef field = (FieldNoteRef)builder.InsertField(FieldType.FieldNoteRef, true);
            field.BookmarkName = bookmarkName;
            field.InsertHyperlink = insertHyperlink;
            field.InsertRelativePosition = insertRelativePosition;
            field.InsertReferenceMark = insertReferenceMark;
            builder.Writeln();
            
            return field;
        }
        
        /// <summary>
        /// Uses a document builder to insert a named bookmark with a footnote at the end.
        /// </summary>
        private static void InsertBookmarkWithFootnote(DocumentBuilder builder, string bookmarkName, string bookmarkText, string footnoteText)
        {
            builder.StartBookmark(bookmarkName);
            builder.Write(bookmarkText);
            builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
            builder.EndBookmark(bookmarkName);
            builder.Writeln();
        }
        //ExEnd

        private void TestNoteRef(Document doc)
        {
            FieldNoteRef field = (FieldNoteRef)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldNoteRef, " NOTEREF  MyBookmark2 \\h", "2", field);
            Assert.AreEqual("MyBookmark2", field.BookmarkName);
            Assert.True(field.InsertHyperlink);
            Assert.False(field.InsertRelativePosition);
            Assert.False(field.InsertReferenceMark);

            field = (FieldNoteRef)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldNoteRef, " NOTEREF  MyBookmark1 \\h \\p", "1 above", field);
            Assert.AreEqual("MyBookmark1", field.BookmarkName);
            Assert.True(field.InsertHyperlink);
            Assert.True(field.InsertRelativePosition);
            Assert.False(field.InsertReferenceMark);

            field = (FieldNoteRef)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldNoteRef, " NOTEREF  MyBookmark2 \\h \\p \\f", "2 below", field);
            Assert.AreEqual("MyBookmark2", field.BookmarkName);
            Assert.True(field.InsertHyperlink);
            Assert.True(field.InsertRelativePosition);
            Assert.True(field.InsertReferenceMark);
        }

        [Test]
        [Ignore("WORDSNET-17845")]
        public void FootnoteRef()
        {
            //ExStart
            //ExFor:FieldFootnoteRef
            //ExSummary:Shows how to cross-reference footnotes with the FOOTNOTEREF field
            // Create a blank document and a document builder for it
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some text, and a footnote, all inside a bookmark named "CrossRefBookmark"
            builder.StartBookmark("CrossRefBookmark");
            builder.Write("Hello world!");
            builder.InsertFootnote(FootnoteType.Footnote, "Cross referenced footnote.");
            builder.EndBookmark("CrossRefBookmark");

            builder.InsertParagraph();
            builder.Write("CrossReference: ");

            // Insert a FOOTNOTEREF field, which lets us reference a footnote more than once while re-using the same footnote marker
            FieldFootnoteRef field = (FieldFootnoteRef) builder.InsertField(FieldType.FieldFootnoteRef, true);

            // Get this field to reference a bookmark
            // The bookmark that we chose contains a footnote marker belonging to the footnote we inserted, which will be displayed by the field, just by itself
            builder.MoveTo(field.Separator);
            builder.Write("CrossRefBookmark");

            Assert.AreEqual(" FOOTNOTEREF CrossRefBookmark", field.GetFieldCode());

            doc.UpdateFields();

            // This field works only in older versions of Microsoft Word
            doc.Save(ArtifactsDir + "Field.FOOTNOTEREF.doc");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.FOOTNOTEREF.doc");
            field = (FieldFootnoteRef)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldFootnoteRef, " FOOTNOTEREF CrossRefBookmark", "1", field);
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, "Cross referenced footnote.", 
                (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
        }

        //ExStart
        //ExFor:FieldPageRef
        //ExFor:FieldPageRef.BookmarkName
        //ExFor:FieldPageRef.InsertHyperlink
        //ExFor:FieldPageRef.InsertRelativePosition
        //ExSummary:Shows to insert PAGEREF fields and present them in different ways.
        [Test] //ExSkip
        [Ignore("WORDSNET-17836")] //ExSkip
        public void FieldPageRef()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            InsertAndNameBookmark(builder, "MyBookmark1");

            // This field will display just the page number where the bookmark starts
            // Setting InsertHyperlink attribute makes the field function as a link to the bookmark
            Assert.AreEqual(" PAGEREF  MyBookmark3 \\h", 
                InsertFieldPageRef(builder, "MyBookmark3", true, false, "Hyperlink to Bookmark3, on page: ").GetFieldCode());

            // Setting the \p flag makes the field display the relative position of the bookmark to the field instead of a page number
            // Bookmark1 is on the same page and above this field, so the result will be "above" on update
            Assert.AreEqual(" PAGEREF  MyBookmark1 \\h \\p", 
                InsertFieldPageRef(builder, "MyBookmark1", true, true, "Bookmark1 is ").GetFieldCode());

            // Bookmark2 will be on the same page and below this field, so the field will display "below"
            Assert.AreEqual(" PAGEREF  MyBookmark2 \\h \\p", 
                InsertFieldPageRef(builder, "MyBookmark2", true, true, "Bookmark2 is ").GetFieldCode());

            // Bookmark3 will be on a different page, so the field will display "on page 2"
            Assert.AreEqual(" PAGEREF  MyBookmark3 \\h \\p", 
                InsertFieldPageRef(builder, "MyBookmark3", true, true, "Bookmark3 is ").GetFieldCode());

            InsertAndNameBookmark(builder, "MyBookmark2");
            builder.InsertBreak(BreakType.PageBreak);
            InsertAndNameBookmark(builder, "MyBookmark3");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.PAGEREF.docx");
            TestPageRef(new Document(ArtifactsDir + "Field.PAGEREF.docx")); //ExSkip
        }

        /// <summary>
        /// Uses a document builder to insert a PAGEREF field and sets its attributes.
        /// </summary>
        private static FieldPageRef InsertFieldPageRef(DocumentBuilder builder, string bookmarkName, bool insertHyperlink, bool insertRelativePosition, string textBefore)
        {
            builder.Write(textBefore);

            FieldPageRef field = (FieldPageRef)builder.InsertField(FieldType.FieldPageRef, true);
            field.BookmarkName = bookmarkName;
            field.InsertHyperlink = insertHyperlink;
            field.InsertRelativePosition = insertRelativePosition;
            builder.Writeln();
          
            return field;
        }

        /// <summary>
        /// Uses a document builder to insert a named bookmark.
        /// </summary>
        private static void InsertAndNameBookmark(DocumentBuilder builder, string bookmarkName)
        {
            builder.StartBookmark(bookmarkName);
            builder.Writeln($"Contents of bookmark \"{bookmarkName}\".");
            builder.EndBookmark(bookmarkName);
        }
        //ExEnd

        private void TestPageRef(Document doc)
        {
            FieldPageRef field = (FieldPageRef)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldPageRef, " PAGEREF  MyBookmark3 \\h", "2", field);
            Assert.AreEqual("MyBookmark3", field.BookmarkName);
            Assert.True(field.InsertHyperlink);
            Assert.False(field.InsertRelativePosition);

            field = (FieldPageRef)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldPageRef, " PAGEREF  MyBookmark1 \\h \\p", "above", field);
            Assert.AreEqual("MyBookmark1", field.BookmarkName);
            Assert.True(field.InsertHyperlink);
            Assert.True(field.InsertRelativePosition);

            field = (FieldPageRef)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldPageRef, " PAGEREF  MyBookmark2 \\h \\p", "below", field);
            Assert.AreEqual("MyBookmark2", field.BookmarkName);
            Assert.True(field.InsertHyperlink);
            Assert.True(field.InsertRelativePosition);

            field = (FieldPageRef)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldPageRef, " PAGEREF  MyBookmark3 \\h \\p", "on page 2", field);
            Assert.AreEqual("MyBookmark3", field.BookmarkName);
            Assert.True(field.InsertHyperlink);
            Assert.True(field.InsertRelativePosition);
        }

        //ExStart
        //ExFor:FieldRef
        //ExFor:FieldRef.BookmarkName
        //ExFor:FieldRef.IncludeNoteOrComment
        //ExFor:FieldRef.InsertHyperlink
        //ExFor:FieldRef.InsertParagraphNumber
        //ExFor:FieldRef.InsertParagraphNumberInFullContext
        //ExFor:FieldRef.InsertParagraphNumberInRelativeContext
        //ExFor:FieldRef.InsertRelativePosition
        //ExFor:FieldRef.NumberSeparator
        //ExFor:FieldRef.SuppressNonDelimiters
        //ExSummary:Shows how to insert REF fields to reference bookmarks and present them in various ways.
        [Test] //ExSkip
        [Ignore("WORDSNET-18067")] //ExSkip
        public void FieldRef()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the bookmark that all our REF fields will reference and leave it at the end of the document
            builder.StartBookmark("MyBookmark");
            builder.InsertFootnote(FootnoteType.Footnote, "MyBookmark footnote #1");
            builder.Write("Text that will appear in REF field");
            builder.InsertFootnote(FootnoteType.Footnote, "MyBookmark footnote #2");
            builder.EndBookmark("MyBookmark");
            builder.MoveToDocumentStart();

            // We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at
            // Note that the angle brackets count as non-delimiter characters
            builder.ListFormat.ApplyNumberDefault();
            builder.ListFormat.ListLevel.NumberFormat = "> \x0000";

            // Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes
            FieldRef field = InsertFieldRef(builder, "MyBookmark", "", "\n");
            field.IncludeNoteOrComment = true;
            field.InsertHyperlink = true;

            Assert.AreEqual(" REF  MyBookmark \\f \\h", field.GetFieldCode());

            // Insert a REF field and display whether the referenced bookmark is above or below it
            field = InsertFieldRef(builder, "MyBookmark", "The referenced paragraph is ", " this field.\n");
            field.InsertRelativePosition = true;

            Assert.AreEqual(" REF  MyBookmark \\p", field.GetFieldCode());

            // Display the list number of the bookmark, as it appears in the document
            field = InsertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number is ", "\n");
            field.InsertParagraphNumber = true;

            Assert.AreEqual(" REF  MyBookmark \\n", field.GetFieldCode());

            // Display the list number of the bookmark, but with non-delimiter characters omitted
            // In this case they are the angle brackets
            field = InsertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number, non-delimiters suppressed, is ", "\n");
            field.InsertParagraphNumber = true;
            field.SuppressNonDelimiters = true;

            Assert.AreEqual(" REF  MyBookmark \\n \\t", field.GetFieldCode());

            // Move down one list level
            builder.ListFormat.ListLevelNumber++;
            builder.ListFormat.ListLevel.NumberFormat = ">> \x0001";

            // Display the list number of the bookmark as well as the numbers of all the list levels above it
            field = InsertFieldRef(builder, "MyBookmark", "The bookmark's full context paragraph number is ", "\n");
            field.InsertParagraphNumberInFullContext = true;

            Assert.AreEqual(" REF  MyBookmark \\w", field.GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);

            // Display the list level numbers between this REF field and the bookmark that it is referencing
            field = InsertFieldRef(builder, "MyBookmark", "The bookmark's relative paragraph number is ", "\n");
            field.InsertParagraphNumberInRelativeContext = true;

            Assert.AreEqual(" REF  MyBookmark \\r", field.GetFieldCode());

            // The bookmark, which is at the end of the document, will show up as a list item here
            builder.Writeln("List level above bookmark");
            builder.ListFormat.ListLevelNumber++;
            builder.ListFormat.ListLevel.NumberFormat = ">>> \x0002";

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.REF.docx");
            TestFieldRef(new Document(ArtifactsDir + "Field.REF.docx")); //ExSkip
        }

        /// <summary>
        /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after.
        /// </summary>
        private static FieldRef InsertFieldRef(DocumentBuilder builder, string bookmarkName, string textBefore, string textAfter)
        {
            builder.Write(textBefore);
            FieldRef field = (FieldRef)builder.InsertField(FieldType.FieldRef, true);
            field.BookmarkName = bookmarkName;
            builder.Write(textAfter);
            return field;
        }
        //ExEnd

        private void TestFieldRef(Document doc)
        {
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, "MyBookmark footnote #1", 
                (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, "MyBookmark footnote #2", 
                (Footnote)doc.GetChild(NodeType.Footnote, 0, true));

            FieldRef field = (FieldRef)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldRef, " REF  MyBookmark \\f \\h", 
                "\u0002 MyBookmark footnote #1\r" +
                "Text that will appear in REF field\u0002 MyBookmark footnote #2\r", field);
            Assert.AreEqual("MyBookmark", field.BookmarkName);
            Assert.True(field.IncludeNoteOrComment);
            Assert.True(field.InsertHyperlink);

            field = (FieldRef)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldRef, " REF  MyBookmark \\p", "below", field);
            Assert.AreEqual("MyBookmark", field.BookmarkName);
            Assert.True(field.InsertRelativePosition);

            field = (FieldRef)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldRef, " REF  MyBookmark \\n", ">>> i", field);
            Assert.AreEqual("MyBookmark", field.BookmarkName);
            Assert.True(field.InsertParagraphNumber);
            Assert.AreEqual(" REF  MyBookmark \\n", field.GetFieldCode());
            Assert.AreEqual(">>> i", field.Result);

            field = (FieldRef)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldRef, " REF  MyBookmark \\n \\t", "i", field);
            Assert.AreEqual("MyBookmark", field.BookmarkName);
            Assert.True(field.InsertParagraphNumber);
            Assert.True(field.SuppressNonDelimiters);

            field = (FieldRef)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldRef, " REF  MyBookmark \\w", "> 4>> c>>> i", field);
            Assert.AreEqual("MyBookmark", field.BookmarkName);
            Assert.True(field.InsertParagraphNumberInFullContext);

            field = (FieldRef)doc.Range.Fields[5];

            TestUtil.VerifyField(FieldType.FieldRef, " REF  MyBookmark \\r", ">> c>>> i", field);
            Assert.AreEqual("MyBookmark", field.BookmarkName);
            Assert.True(field.InsertParagraphNumberInRelativeContext);
        }

        [Test]
        [Ignore("WORDSNET-18068")]
        public void FieldRD()
        {
            //ExStart
            //ExFor:FieldRD
            //ExFor:FieldRD.FileName
            //ExFor:FieldRD.IsPathRelative
            //ExSummary:Shows to insert an RD field to source table of contents entries from an external document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a table of contents and, on the following page, one entry
            builder.InsertField(FieldType.FieldTOC, true);
            builder.InsertBreak(BreakType.PageBreak);
            builder.CurrentParagraph.ParagraphFormat.StyleName = "Heading 1";
            builder.Writeln("TOC entry from within this document");

            // Insert an RD field, designating an external document that our TOC field will look in for more entries
            FieldRD field = (FieldRD)builder.InsertField(FieldType.FieldRefDoc, true);
            field.FileName = "ReferencedDocument.docx";
            field.IsPathRelative = true;
            field.Update();

            Assert.AreEqual(" RD  ReferencedDocument.docx \\f", field.GetFieldCode());

            // Create the document and insert a TOC entry, which will end up in the TOC of our original document
            Document referencedDoc = new Document();
            DocumentBuilder refDocBuilder = new DocumentBuilder(referencedDoc);
            refDocBuilder.CurrentParagraph.ParagraphFormat.StyleName = "Heading 1";
            refDocBuilder.Writeln("TOC entry from referenced document");
            referencedDoc.Save(ArtifactsDir + "ReferencedDocument.docx");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.RD.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.RD.docx");

            FieldToc fieldToc = (FieldToc)doc.Range.Fields[0];

            Assert.AreEqual("TOC entry from within this document\t\u0013 PAGEREF _Toc36149519 \\h \u00142\u0015\r" +
                            "TOC entry from referenced document\t1\r", fieldToc.Result);

            FieldPageRef fieldPageRef = (FieldPageRef)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldPageRef, " PAGEREF _Toc36149519 \\h ", "2", fieldPageRef);

            field = (FieldRD)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldRefDoc, " RD  ReferencedDocument.docx \\f", string.Empty, field);
            Assert.AreEqual("ReferencedDocument.docx", field.FileName);
            Assert.True(field.IsPathRelative);
        }

        [Test]
        public void SkipIf()
        {
            //ExStart
            //ExFor:FieldSkipIf
            //ExFor:FieldSkipIf.ComparisonOperator
            //ExFor:FieldSkipIf.LeftExpression
            //ExFor:FieldSkipIf.RightExpression
            //ExSummary:Shows how to skip pages in a mail merge using the SKIPIF field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a data table that will be the source for our mail merge
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Name");
            table.Columns.Add("Department");
            table.Rows.Add("John Doe", "Sales");
            table.Rows.Add("Jane Doe", "Accounting");
            table.Rows.Add("John Cardholder", "HR");

            // Insert a SKIPIF field, which will skip a page of a mail merge if the condition is fulfilled
            // We will move to the SKIPIF field's separator character and insert a MERGEFIELD at that place to create a nested field
            FieldSkipIf fieldSkipIf = (FieldSkipIf) builder.InsertField(FieldType.FieldSkipIf, true);
            builder.MoveTo(fieldSkipIf.Separator);
            FieldMergeField fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Department";

            // The MERGEFIELD refers to the "Department" column in our data table, and our SKIPIF field will check if its value equals to "HR"
            // One of three rows satisfy that condition, so we will expect the result of our mail merge to have two pages
            fieldSkipIf.LeftExpression = "=";
            fieldSkipIf.RightExpression = "HR";

            // Add some content to our mail merge and execute it
            builder.MoveToDocumentEnd();
            builder.Write("Dear ");
            fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Name";
            builder.Writeln(", ");

            doc.MailMerge.Execute(table);
            doc.Save(ArtifactsDir + "Field.SKIPIF.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SKIPIF.docx");

            Assert.AreEqual(0, doc.Range.Fields.Count);
            Assert.AreEqual("Dear John Doe, \r" +
                            "\fDear Jane Doe, \r\f", doc.GetText());
        }
      
        [Test]
        public void FieldSet()
        {
            //ExStart
            //ExFor:FieldSet
            //ExFor:FieldSet.BookmarkName
            //ExFor:FieldSet.BookmarkText
            //ExSummary:Shows to alter a bookmark's text with a SET field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("MyBookmark");
            builder.Writeln("Bookmark contents");
            builder.EndBookmark("MyBookmark");

            Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];
            bookmark.Text = "Old text";

            FieldSet field = (FieldSet)builder.InsertField(FieldType.FieldSet, false);
            field.BookmarkName = "MyBookmark";
            field.BookmarkText = "New text";

            Assert.AreEqual(" SET  MyBookmark \"New text\"", field.GetFieldCode());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.SET.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SET.docx");

            Assert.AreEqual("New text", doc.Range.Bookmarks[0].Text);

            field = (FieldSet)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldSet, " SET  MyBookmark \"New text\"", "New text", field);
            Assert.AreEqual("MyBookmark", field.BookmarkName);
            Assert.AreEqual("New text", field.BookmarkText);
        }

        [Test]
        [Ignore("WORDSNET-18137")]
        public void FieldTemplate()
        {
            //ExStart
            //ExFor:FieldTemplate
            //ExFor:FieldTemplate.IncludeFullPath
            //ExSummary:Shows how to display the location of the document's template with a TEMPLATE field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldTemplate field = (FieldTemplate)builder.InsertField(FieldType.FieldTemplate, false);
            Assert.AreEqual(" TEMPLATE ", field.GetFieldCode());

            builder.Writeln();
            field = (FieldTemplate)builder.InsertField(FieldType.FieldTemplate, false);
            field.IncludeFullPath = true;

            Assert.AreEqual(" TEMPLATE  \\p", field.GetFieldCode());

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.TEMPLATE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.TEMPLATE.docx");

            field = (FieldTemplate)doc.Range.Fields[0];
            Assert.AreEqual(" TEMPLATE ", field.GetFieldCode());
            Assert.AreEqual("Normal.dotm", field.Result);

            field = (FieldTemplate)doc.Range.Fields[1];
            Assert.AreEqual(" TEMPLATE  \\p", field.GetFieldCode());
            Assert.True(field.Result.EndsWith("\\Microsoft\\Templates\\Normal.dotm"));

        }

        [Test]
        public void FieldSymbol()
        {
            //ExStart
            //ExFor:FieldSymbol
            //ExFor:FieldSymbol.CharacterCode
            //ExFor:FieldSymbol.DontAffectsLineSpacing
            //ExFor:FieldSymbol.FontName
            //ExFor:FieldSymbol.FontSize
            //ExFor:FieldSymbol.IsAnsi
            //ExFor:FieldSymbol.IsShiftJis
            //ExFor:FieldSymbol.IsUnicode
            //ExSummary:Shows how to use the SYMBOL field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a SYMBOL field to display a symbol, designated by a character code
            FieldSymbol field = (FieldSymbol)builder.InsertField(FieldType.FieldSymbol, true);

            // The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol 
            field.CharacterCode = 0x00a9.ToString();
            field.IsAnsi = true;

            Assert.AreEqual(" SYMBOL  169 \\a", field.GetFieldCode());

            builder.Writeln(" Line 1");

            // In Unicode, the "221E" code is reserved for the infinity symbol
            field = (FieldSymbol)builder.InsertField(FieldType.FieldSymbol, true);
            field.CharacterCode = 0x221E.ToString();
            field.IsUnicode = true;

            // Change the appearance of our symbol
            // Note that some symbols can change from font to font
            // The full list of symbols and their fonts can be looked up in the Windows Character Map
            field.FontName = "Calibri";
            field.FontSize = "24";

            // A tall symbol like the one we placed can also be made to not push down the text on its line
            field.DontAffectsLineSpacing = true;

            Assert.AreEqual(" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", field.GetFieldCode());

            builder.Writeln("Line 2");

            // Display a symbol from the Shift-JIS, also known as the Windows-932 code page
            // With a font that supports Shift-JIS, this symbol will display "あ"
            field = (FieldSymbol)builder.InsertField(FieldType.FieldSymbol, true);
            field.FontName = "MS Gothic";
            field.CharacterCode = 0x82A0.ToString();
            field.IsShiftJis = true;

            Assert.AreEqual(" SYMBOL  33440 \\f \"MS Gothic\" \\j", field.GetFieldCode());

            builder.Write("Line 3");

            doc.Save(ArtifactsDir + "Field.SYMBOL.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SYMBOL.docx");

            field = (FieldSymbol)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldSymbol, " SYMBOL  169 \\a", string.Empty, field);
            Assert.AreEqual(0x00a9.ToString(), field.CharacterCode);
            Assert.True(field.IsAnsi);
            Assert.AreEqual("©", field.DisplayResult);
                
            field = (FieldSymbol)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldSymbol, " SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", string.Empty, field);
            Assert.AreEqual(0x221E.ToString(), field.CharacterCode);
            Assert.AreEqual("Calibri", field.FontName);
            Assert.AreEqual("24", field.FontSize);
            Assert.True(field.IsUnicode);
            Assert.True(field.DontAffectsLineSpacing);
            Assert.AreEqual("∞", field.DisplayResult);

            field = (FieldSymbol)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldSymbol, " SYMBOL  33440 \\f \"MS Gothic\" \\j", string.Empty, field);
            Assert.AreEqual(0x82A0.ToString(), field.CharacterCode);
            Assert.AreEqual("MS Gothic", field.FontName);
            Assert.True(field.IsShiftJis);
        }

        [Test]
        public void FieldTitle()
        {
            //ExStart
            //ExFor:FieldTitle
            //ExFor:FieldTitle.Text
            //ExSummary:Shows how to use the TITLE field.
            Document doc = new Document();

            // A TITLE field will display the value assigned to this variable
            doc.BuiltInDocumentProperties.Title = "My Title";

            // Insert a TITLE field using a document builder
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldTitle field = (FieldTitle)builder.InsertField(FieldType.FieldTitle, false);
            field.Update();

            Assert.AreEqual(" TITLE ", field.GetFieldCode());
            Assert.AreEqual("My Title", field.Result);

            // Set the Text attribute to display a different value
            builder.Writeln();
            field = (FieldTitle)builder.InsertField(FieldType.FieldTitle, false);
            field.Text = "My New Title";
            field.Update();

            Assert.AreEqual(" TITLE  \"My New Title\"", field.GetFieldCode());
            Assert.AreEqual("My New Title", field.Result);

            // In doing that we have also changed the title in the document properties
            Assert.AreEqual("My New Title", doc.BuiltInDocumentProperties.Title);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.TITLE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.TITLE.docx");

            Assert.AreEqual("My New Title", doc.BuiltInDocumentProperties.Title);

            field = (FieldTitle)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldTitle, " TITLE ", "My New Title", field);

            field = (FieldTitle)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldTitle, " TITLE  \"My New Title\"", "My New Title", field);
            Assert.AreEqual("My New Title", field.Text);
        }

        //ExStart
        //ExFor:FieldToa
        //ExFor:FieldToa.BookmarkName
        //ExFor:FieldToa.EntryCategory
        //ExFor:FieldToa.EntrySeparator
        //ExFor:FieldToa.PageNumberListSeparator
        //ExFor:FieldToa.PageRangeSeparator
        //ExFor:FieldToa.RemoveEntryFormatting
        //ExFor:FieldToa.SequenceName
        //ExFor:FieldToa.SequenceSeparator
        //ExFor:FieldToa.UseHeading
        //ExFor:FieldToa.UsePassim
        //ExFor:FieldTA
        //ExFor:FieldTA.EntryCategory
        //ExFor:FieldTA.IsBold
        //ExFor:FieldTA.IsItalic
        //ExFor:FieldTA.LongCitation
        //ExFor:FieldTA.PageRangeBookmarkName
        //ExFor:FieldTA.ShortCitation
        //ExSummary:Shows how to build and customize a table of authorities using TOA and TA fields.
        [Test] //ExSkip
        public void FieldTOA()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TOA field, which will list all the TA entries in the document,
            // displaying long citations and page numbers for each
            FieldToa fieldToa = (FieldToa)builder.InsertField(FieldType.FieldTOA, false);

            // Set the entry category for our table
            // For a TA field to be included in this table, it will have to have a matching entry category
            fieldToa.EntryCategory = "1";

            // Moreover, the Table of Authorities category at index 1 is "Cases",
            // which will show up as the title of our table if we set this variable to true
            fieldToa.UseHeading = true;

            // We can further filter TA fields by designating a named bookmark that they have to be inside of
            fieldToa.BookmarkName = "MyBookmark";

            // By default, a dotted line page-wide tab appears between the TA field's citation and its page number
            // We can replace it with any text we put in this attribute, even preserving the tab if we use tab character
            fieldToa.EntrySeparator = " \t p.";

            // If we have multiple TA entries that share the same long citation,
            // all their respective page numbers will show up on one row,
            // and the page numbers separated by a string specified here
            fieldToa.PageNumberListSeparator = " & p. ";

            // To reduce clutter, we can set this to true to get our table to display the word "passim"
            // if there are 5 or more page numbers in one row
            fieldToa.UsePassim = true;

            // One TA field can refer to a range of pages, and the sequence specified here will be between the start and end page numbers
            fieldToa.PageRangeSeparator = " to ";

            // The format from the TA fields will carry over into our table, and we can stop it from doing so by setting this variable
            fieldToa.RemoveEntryFormatting = true;
            builder.Font.Color = Color.Green;
            builder.Font.Name = "Arial Black";

            Assert.AreEqual(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldToa.GetFieldCode());

            builder.InsertBreak(BreakType.PageBreak);

            // We will insert a TA entry using a document builder
            // This entry is outside the bookmark specified by our table, so it will not be displayed
            FieldTA fieldTA = InsertToaEntry(builder, "1", "Source 1");

            Assert.AreEqual(" TA  \\c 1 \\l \"Source 1\"", fieldTA.GetFieldCode());

            // This entry is inside the bookmark,
            // but the entry category does not match that of the table, so it will also be omitted
            builder.StartBookmark("MyBookmark");
            fieldTA = InsertToaEntry(builder, "2", "Source 2");

            // This entry will appear in the table
            fieldTA = InsertToaEntry(builder, "1", "Source 3");

            // Short citations are not displayed by a TOA table,
            // but they can be used as a shorthand to refer to bulky source names that multiple TA fields reference
            fieldTA.ShortCitation = "S.3";

            Assert.AreEqual(" TA  \\c 1 \\l \"Source 3\" \\s S.3", fieldTA.GetFieldCode());

            // The page number can be made to appear bold and/or italic
            // This will still be displayed if our table is set to ignore formatting
            fieldTA = InsertToaEntry(builder, "1", "Source 2");
            fieldTA.IsBold = true;
            fieldTA.IsItalic = true;

            Assert.AreEqual(" TA  \\c 1 \\l \"Source 2\" \\b \\i", fieldTA.GetFieldCode());

            // We can get TA fields to refer to a range of pages that a bookmark spans across instead of the page that they are on
            // Note that this entry refers to the same source as the one above, so they will share one row in our table,
            // displaying the page number of the entry above as well as the page range of this entry,
            // with the table's page list and page number range separators between page numbers
            fieldTA = InsertToaEntry(builder, "1", "Source 3");
            fieldTA.PageRangeBookmarkName = "MyMultiPageBookmark";

            builder.StartBookmark("MyMultiPageBookmark");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);
            builder.EndBookmark("MyMultiPageBookmark");

            Assert.AreEqual(" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", fieldTA.GetFieldCode());

            // Having 5 or more TA entries with the same source invokes the "passim" feature of our table, if we enabled it
            for (int i = 0; i < 5; i++)
            {
                InsertToaEntry(builder, "1", "Source 4");
            }

            builder.EndBookmark("MyBookmark");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.TOA.TA.docx");
            TestFieldTOA(new Document(ArtifactsDir + "Field.TOA.TA.docx")); //ExSKip
        }

        /// <summary>
        /// Get a builder to insert a TA field, specifying its long citation and category,
        /// then insert a page break and return the field we created.
        /// </summary>
        private static FieldTA InsertToaEntry(DocumentBuilder builder, string entryCategory, string longCitation)
        {
            FieldTA field = (FieldTA)builder.InsertField(FieldType.FieldTOAEntry, false);
            field.EntryCategory = entryCategory;
            field.LongCitation = longCitation;

            builder.InsertBreak(BreakType.PageBreak);

            return field;
        }
        //ExEnd

        private void TestFieldTOA(Document doc)
        {
            FieldToa fieldTOA = (FieldToa)doc.Range.Fields[0];

            Assert.AreEqual("1", fieldTOA.EntryCategory);
            Assert.True(fieldTOA.UseHeading);
            Assert.AreEqual("MyBookmark", fieldTOA.BookmarkName);
            Assert.AreEqual(" \t p.", fieldTOA.EntrySeparator);
            Assert.AreEqual(" & p. ", fieldTOA.PageNumberListSeparator);
            Assert.True(fieldTOA.UsePassim);
            Assert.AreEqual(" to ", fieldTOA.PageRangeSeparator);
            Assert.True(fieldTOA.RemoveEntryFormatting);
            Assert.AreEqual(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldTOA.GetFieldCode());
            Assert.AreEqual("Cases\r" +
                            "Source 2 \t p.5\r" +
                            "Source 3 \t p.4 & p. 7 to 10\r" +
                            "Source 4 \t p.passim\r", fieldTOA.Result);

            FieldTA fieldTA = (FieldTA)doc.Range.Fields[1];

            TestUtil.VerifyField(FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 1\"", string.Empty, fieldTA);
            Assert.AreEqual("1", fieldTA.EntryCategory);
            Assert.AreEqual("Source 1", fieldTA.LongCitation);

            fieldTA = (FieldTA)doc.Range.Fields[2];

            TestUtil.VerifyField(FieldType.FieldTOAEntry, " TA  \\c 2 \\l \"Source 2\"", string.Empty, fieldTA);
            Assert.AreEqual("2", fieldTA.EntryCategory);
            Assert.AreEqual("Source 2", fieldTA.LongCitation);

            fieldTA = (FieldTA)doc.Range.Fields[3];

            TestUtil.VerifyField(FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 3\" \\s S.3", string.Empty, fieldTA);
            Assert.AreEqual("1", fieldTA.EntryCategory);
            Assert.AreEqual("Source 3", fieldTA.LongCitation);
            Assert.AreEqual("S.3", fieldTA.ShortCitation);

            fieldTA = (FieldTA)doc.Range.Fields[4];

            TestUtil.VerifyField(FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 2\" \\b \\i", string.Empty, fieldTA);
            Assert.AreEqual("1", fieldTA.EntryCategory);
            Assert.AreEqual("Source 2", fieldTA.LongCitation);
            Assert.True(fieldTA.IsBold);
            Assert.True(fieldTA.IsItalic);

            fieldTA = (FieldTA)doc.Range.Fields[5];

            TestUtil.VerifyField(FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", string.Empty, fieldTA);
            Assert.AreEqual("1", fieldTA.EntryCategory);
            Assert.AreEqual("Source 3", fieldTA.LongCitation);
            Assert.AreEqual("MyMultiPageBookmark", fieldTA.PageRangeBookmarkName);

            for (int i = 6; i < 11; i++)
            {
                fieldTA = (FieldTA)doc.Range.Fields[i];

                TestUtil.VerifyField(FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 4\"", string.Empty, fieldTA);
                Assert.AreEqual("1", fieldTA.EntryCategory);
                Assert.AreEqual("Source 4", fieldTA.LongCitation);
            }
        }

        [Test]
        public void FieldAddIn()
        {
            //ExStart
            //ExFor:FieldAddIn
            //ExSummary:Shows how to process an ADDIN field.
            // Open a document that contains an ADDIN field
            Document doc = new Document(MyDir + "Field sample - ADDIN.docx");

            // Aspose.Words does not support inserting ADDIN fields, they can be read
            FieldAddIn field = (FieldAddIn)doc.Range.Fields[0];

            Assert.AreEqual(" ADDIN \"My value\" ", field.GetFieldCode());
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            TestUtil.VerifyField(FieldType.FieldAddin, " ADDIN \"My value\" ", string.Empty, doc.Range.Fields[0]);
        }

        [Test]
        public void FieldEditTime()
        {
            //ExStart
            //ExFor:FieldEditTime
            //ExSummary:Shows how to use the EDITTIME field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert an EDITTIME field in the header
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("You've been editing this document for ");
            FieldEditTime field = (FieldEditTime)builder.InsertField(FieldType.FieldEditTime, true);
            builder.Writeln(" minutes.");

            // The EDITTIME field will show, in minutes only,
            // the time spent with the document open in a Microsoft Word window
            // The minutes are tracked in a document property, which we can change like this
            doc.BuiltInDocumentProperties.TotalEditingTime = 10;
            field.Update();

            Assert.AreEqual(" EDITTIME ", field.GetFieldCode());
            Assert.AreEqual("10", field.Result);

            // The field does not update in real time and will have to be manually updated in Microsoft Word also
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.EDITTIME.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.EDITTIME.docx");

            Assert.AreEqual(10, doc.BuiltInDocumentProperties.TotalEditingTime);

            TestUtil.VerifyField(FieldType.FieldEditTime, " EDITTIME ", "10", doc.Range.Fields[0]);
        }

        //ExStart
        //ExFor:FieldEQ
        //ExSummary:Shows how to use the EQ field to display a variety of mathematical equations.
        [Test] //ExSkip
        public void FieldEQ()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // An EQ field displays a mathematical equation consisting of one or many elements
            // Each element takes the following form: [switch][options][arguments]
            // One switch, several possible options, followed by a set of argument values inside round braces

            // Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction"
            // No options are invoked, and the values 1 and 4 are passed as arguments
            // This field will display a fraction with 1 as the numerator and 4 as the denominator
            FieldEQ field = InsertFieldEQ(builder, @"\f(1,4)");

            Assert.AreEqual(@" EQ \f(1,4)", field.GetFieldCode());

            // One EQ field may contain multiple elements placed sequentially,
            // and elements may also be nested by being placed inside the argument brackets of other elements
            // The full list of switches and their corresponding options can be found here:
            // https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/

            // Array switch "\a", aligned left, 2 columns, 3 points of horizontal and vertical spacing
            InsertFieldEQ(builder, @"\a \al \co2 \vs3 \hs3(4x,- 4y,-4x,+ y)");

            // Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces
            // Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output
            InsertFieldEQ(builder, @"\b \bc\[ (\a \al \co3 \vs3 \hs3(1,0,0,0,1,0,0,0,1))");

            // Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline
            InsertFieldEQ(builder, @"A \d \fo30 \li() B");

            // Formula consisting of multiple fractions
            InsertFieldEQ(builder, @"\f(d,dx)(u + v) = \f(du,dx) + \f(dv,dx)");

            // Integral switch "\i", with a summation symbol
            InsertFieldEQ(builder, @"\i \su(n=1,5,n)");

            // List switch "\l"
            InsertFieldEQ(builder, @"\l(1,1,2,3,n,8,13)");

            // Radical switch "\r", displaying a cubed root of x
            InsertFieldEQ(builder, @"\r (3,x)");

            // Subscript/superscript switch "/s", first as a superscript and then as a subscript
            InsertFieldEQ(builder, @"\s \up8(Superscript) Text \s \do8(Subscript)");

            // Box switch "\x", with lines at the top, bottom, left and right of the input
            InsertFieldEQ(builder, @"\x \to \bo \le \ri(5)");

            // More complex combinations
            InsertFieldEQ(builder, @"\a \ac \vs1 \co1(lim,n→∞) \b (\f(n,n2 + 12) + \f(n,n2 + 22) + ... + \f(n,n2 + n2))");
            InsertFieldEQ(builder, @"\i (,,  \b(\f(x,x2 + 3x + 2))) \s \up10(2)");
            InsertFieldEQ(builder, @"\i \in( tan x, \s \up2(sec x), \b(\r(3) )\s \up4(t) \s \up7(2)  dt)");

            doc.Save(ArtifactsDir + "Field.EQ.docx");
            TestFieldEQ(new Document(ArtifactsDir + "Field.EQ.docx")); //ExSkip
        }

        /// <summary>
        /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
        /// </summary>
        private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
        {
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            builder.MoveTo(field.Separator);
            builder.Write(args);
            builder.MoveTo(field.Start.ParentNode);
            
            builder.InsertParagraph();
            return field;
        }
        //ExEnd

        private void TestFieldEQ(Document doc)
        {
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \f(1,4)", string.Empty, doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \a \al \co2 \vs3 \hs3(4x,- 4y,-4x,+ y)", string.Empty, doc.Range.Fields[1]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \b \bc\[ (\a \al \co3 \vs3 \hs3(1,0,0,0,1,0,0,0,1))", string.Empty, doc.Range.Fields[2]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ A \d \fo30 \li() B", string.Empty, doc.Range.Fields[3]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \f(d,dx)(u + v) = \f(du,dx) + \f(dv,dx)", string.Empty, doc.Range.Fields[4]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \i \su(n=1,5,n)", string.Empty, doc.Range.Fields[5]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \l(1,1,2,3,n,8,13)", string.Empty, doc.Range.Fields[6]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \r (3,x)", string.Empty, doc.Range.Fields[7]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \s \up8(Superscript) Text \s \do8(Subscript)", string.Empty, doc.Range.Fields[8]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \x \to \bo \le \ri(5)", string.Empty, doc.Range.Fields[9]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \a \ac \vs1 \co1(lim,n→∞) \b (\f(n,n2 + 12) + \f(n,n2 + 22) + ... + \f(n,n2 + n2))", string.Empty, doc.Range.Fields[10]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \i (,,  \b(\f(x,x2 + 3x + 2))) \s \up10(2)", string.Empty, doc.Range.Fields[11]);
            TestUtil.VerifyField(FieldType.FieldEquation, @" EQ \i \in( tan x, \s \up2(sec x), \b(\r(3) )\s \up4(t) \s \up7(2)  dt)", string.Empty, doc.Range.Fields[12]);
        }

        [Test]
        public void FieldForms()
        {
            //ExStart
            //ExFor:FieldFormCheckBox
            //ExFor:FieldFormDropDown
            //ExFor:FieldFormText
            //ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
            // These fields are legacy equivalents of the FormField, and they can be read but not inserted by Aspose.Words,
            // and can be inserted in Microsoft Word 2019 via the Legacy Tools menu in the Developer tab
            Document doc = new Document(MyDir + "Form fields.docx");

            FieldFormCheckBox fieldFormCheckBox = (FieldFormCheckBox)doc.Range.Fields[1];
            Assert.AreEqual(" FORMCHECKBOX \u0001", fieldFormCheckBox.GetFieldCode());

            FieldFormDropDown fieldFormDropDown = (FieldFormDropDown)doc.Range.Fields[2];
            Assert.AreEqual(" FORMDROPDOWN \u0001", fieldFormDropDown.GetFieldCode());

            FieldFormText fieldFormText = (FieldFormText)doc.Range.Fields[0];
            Assert.AreEqual(" FORMTEXT \u0001", fieldFormText.GetFieldCode());
            //ExEnd
        }

        [Test]
        public void FieldFormula()
        {
            //ExStart
            //ExFor:FieldFormula
            //ExSummary:Shows how to use the "=" field.
            Document doc = new Document();

            // Create a formula field using a field builder
            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldFormula);
            fieldBuilder.AddArgument(2);
            fieldBuilder.AddArgument("*");
            fieldBuilder.AddArgument(5);

            FieldFormula field = (FieldFormula)fieldBuilder.BuildAndInsert(doc.FirstSection.Body.FirstParagraph);
            field.Update();

            Assert.AreEqual(" = 2 * 5 ", field.GetFieldCode());
            Assert.AreEqual("10", field.Result);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.FORMULA.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.FORMULA.docx");

            TestUtil.VerifyField(FieldType.FieldFormula, " = 2 * 5 ", "10", doc.Range.Fields[0]);
        }

        [Test]
        public void FieldLastSavedBy()
        {
            //ExStart
            //ExFor:FieldLastSavedBy
            //ExSummary:Shows how to use the LASTSAVEDBY field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" property
            // This is the property that a LASTSAVEDBY field looks up and displays
            // If we make a document programmatically, this property is null and needs to have a value assigned to it first
            doc.BuiltInDocumentProperties.LastSavedBy = "John Doe";

            // Insert a LASTSAVEDBY field using a document builder
            FieldLastSavedBy field = (FieldLastSavedBy)builder.InsertField(FieldType.FieldLastSavedBy, true);

            // The value from our document property appears here
            Assert.AreEqual(" LASTSAVEDBY ", field.GetFieldCode());
            Assert.AreEqual("John Doe", field.Result);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.LASTSAVEDBY.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.LASTSAVEDBY.docx");

            Assert.AreEqual("John Doe", doc.BuiltInDocumentProperties.LastSavedBy);
            TestUtil.VerifyField(FieldType.FieldLastSavedBy, " LASTSAVEDBY ", "John Doe", doc.Range.Fields[0]);
        }

        [Test]
        [Ignore("WORDSNET-18173")]
        public void FieldMergeRec()
        {
            //ExStart
            //ExFor:FieldMergeRec
            //ExFor:FieldMergeSeq
            //ExSummary:Shows how to number and count mail merge records in the output documents of a mail merge using MERGEREC and MERGESEQ fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a merge field
            builder.Write("Dear ");
            FieldMergeField fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Name";
            builder.Writeln(",");

            // A MERGEREC field will print the row number of the data being merged
            builder.Write("\nRow number of record in data source: ");
            FieldMergeRec fieldMergeRec = (FieldMergeRec)builder.InsertField(FieldType.FieldMergeRec, true);

            Assert.AreEqual(" MERGEREC ", fieldMergeRec.GetFieldCode());

            // A MERGESEQ field will count the number of successful merges and print the current value on each respective page
            // If no rows are skipped and the data source is not sorted, and no SKIP/SKIPIF/NEXT/NEXTIF fields are invoked,
            // the MERGESEQ and MERGEREC fields will function the same
            builder.Write("\nSuccessful merge number: ");
            FieldMergeSeq fieldMergeSeq = (FieldMergeSeq)builder.InsertField(FieldType.FieldMergeSeq, true);

            Assert.AreEqual(" MERGESEQ ", fieldMergeSeq.GetFieldCode());

            // Insert a SKIPIF field, which will skip a merge if the name is "John Doe"
            FieldSkipIf fieldSkipIf = (FieldSkipIf)builder.InsertField(FieldType.FieldSkipIf, true);
            builder.MoveTo(fieldSkipIf.Separator);
            fieldMergeField = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, true);
            fieldMergeField.FieldName = "Name";
            fieldSkipIf.LeftExpression = "=";
            fieldSkipIf.RightExpression = "John Doe";

            // Create a data source with 3 rows, one of them having "John Doe" as a value for the "Name" column
            // Since a SKIPIF field will be triggered once by that value, the output of our mail merge will have 2 pages instead of 3
            // On page 1, the MERGESEQ and MERGEREC fields will both display "1"
            // On page 2, the MERGEREC field will display "3" and the MERGESEQ field will display "2"
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Name");
            table.Rows.Add(new[] { "Jane Doe" });
            table.Rows.Add(new[] { "John Doe" });
            table.Rows.Add(new[] { "Joe Bloggs" });

            // Execute the mail merge and save document
            doc.MailMerge.Execute(table);
            doc.Save(ArtifactsDir + "Field.MERGEREC.MERGESEQ.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.MERGEREC.MERGESEQ.docx");

            Assert.AreEqual(0, doc.Range.Fields.Count);

            Assert.AreEqual("Dear Jane Doe,\r" +
                            "\r" +
                            "Row number of record in data source: 1\r" +
                            "Successful merge number: 1\fDear Joe Bloggs,\r" +
                            "\r" +
                            "Row number of record in data source: 2\r" +
                            "Successful merge number: 3", doc.GetText().Trim());
        }

        [Test]
        public void FieldOcx()
        {
            //ExStart
            //ExFor:FieldOcx
            //ExSummary:Shows how to insert an OCX field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert an OCX field
            FieldOcx field = (FieldOcx)builder.InsertField(FieldType.FieldOcx, true);

            Assert.AreEqual(" OCX ", field.GetFieldCode());
            //ExEnd

            TestUtil.VerifyField(FieldType.FieldOcx, " OCX ", string.Empty, field);
        }

        //ExStart
        //ExFor:Field.Remove
        //ExFor:FieldPrivate
        //ExSummary:Shows how to process PRIVATE fields.
        [Test] //ExSkip
        public void FieldPrivate()
        {
            // Open a Corel WordPerfect document that was converted to .docx format
            Document doc = new Document(MyDir + "Field sample - PRIVATE.docx");

            // WordPerfect 5.x/6.x documents like the one we opened may contain PRIVATE fields
            // The PRIVATE field is a WordPerfect artifact that is preserved when a file is opened and saved in Microsoft Word
            // However, they have no functionality in Microsoft Word
            FieldPrivate field = (FieldPrivate)doc.Range.Fields[0];

            Assert.AreEqual(" PRIVATE \"My value\" ", field.GetFieldCode());
            Assert.AreEqual(FieldType.FieldPrivate, field.Type);

            // PRIVATE fields can also be inserted by a document builder
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(FieldType.FieldPrivate, true);

            // It is strongly advised against using them to attempt to hide or store private information
            // Unless backward compatibility with older versions of WordPerfect is necessary, these fields can safely be removed
            // This can be done using a document visitor implementation
            Assert.AreEqual(2, doc.Range.Fields.Count);

            FieldPrivateRemover remover = new FieldPrivateRemover();
            doc.Accept(remover);

            Assert.AreEqual(2, remover.GetFieldsRemovedCount());
            Assert.AreEqual(0, doc.Range.Fields.Count);
        }

        /// <summary>
        /// Visitor implementation that removes all PRIVATE fields that it encounters.
        /// </summary>
        public class FieldPrivateRemover : DocumentVisitor
        {
            public FieldPrivateRemover()
            {
                mFieldsRemovedCount = 0;
            }

            public int GetFieldsRemovedCount()
            {
                return mFieldsRemovedCount;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// If the node belongs to a PRIVATE field, the entire field is removed.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                if (fieldEnd.FieldType == FieldType.FieldPrivate)
                {
                    fieldEnd.GetField().Remove();
                    mFieldsRemovedCount++;
                }

                return VisitorAction.Continue;
            }

            private int mFieldsRemovedCount;
        }
        //ExEnd

        [Test]
        public void FieldSection()
        {
            //ExStart
            //ExFor:FieldSection
            //ExFor:FieldSectionPages
            //ExSummary:Shows how to use SECTION and SECTIONPAGES fields to facilitate page numbering separated by sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the document builder to the header that appears across all pages and align to the top right
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            // A SECTION field displays the number of the section it is placed in
            builder.Write("Section ");
            FieldSection fieldSection = (FieldSection)builder.InsertField(FieldType.FieldSection, true);

            Assert.AreEqual(" SECTION ", fieldSection.GetFieldCode());

            // A PAGE field displays the number of the page it is placed in
            builder.Write("\nPage ");
            FieldPage fieldPage = (FieldPage)builder.InsertField(FieldType.FieldPage, true);

            Assert.AreEqual(" PAGE ", fieldPage.GetFieldCode());

            // A SECTIONPAGES field displays the number of pages that the section it is in spans across
            builder.Write(" of ");
            FieldSectionPages fieldSectionPages = (FieldSectionPages)builder.InsertField(FieldType.FieldSectionPages, true);

            Assert.AreEqual(" SECTIONPAGES ", fieldSectionPages.GetFieldCode());

            // Move out of the header back into the main document and insert two pages
            // Both these pages will be in the first section and our three fields will keep track of the numbers in each header
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);

            // We can insert a new section with the document builder like this
            // This will change the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // The PAGE field will keep counting pages across the whole document
            // We can manually reset its count after a new section is added to keep track of pages section-by-section
            builder.CurrentSection.PageSetup.RestartPageNumbering = true;
            builder.InsertBreak(BreakType.PageBreak);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "Field.SECTION.SECTIONPAGES.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.SECTION.SECTIONPAGES.docx");

            TestUtil.VerifyField(FieldType.FieldSection, " SECTION ", "2", doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldPage, " PAGE ", "2", doc.Range.Fields[1]);
            TestUtil.VerifyField(FieldType.FieldSectionPages, " SECTIONPAGES ", "2", doc.Range.Fields[2]);
        }

        //ExStart
        //ExFor:FieldTime
        //ExSummary:Shows how to display the current time using the TIME field.
        [Test] //ExSkip
        public void FieldTime()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default, time is displayed in the "h:mm am/pm" format
            FieldTime field = InsertFieldTime(builder, "");

            Assert.AreEqual(" TIME ", field.GetFieldCode());

            // By using the \@ flag, we can change the appearance of our time
            field = InsertFieldTime(builder, "\\@ HHmm");

            Assert.AreEqual(" TIME \\@ HHmm", field.GetFieldCode());

            // We can even display the date, according to the Gregorian calendar
            field = InsertFieldTime(builder, "\\@ \"M/d/yyyy h mm:ss am/pm\"");

            Assert.AreEqual(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.GetFieldCode());

            doc.Save(ArtifactsDir + "Field.TIME.docx");
            TestFieldTime(new Document(ArtifactsDir + "Field.TIME.docx")); //ExSkip
        }

        /// <summary>
        /// Use a document builder to insert a TIME field, insert a new paragraph and return the field
        /// </summary>
        private static FieldTime InsertFieldTime(DocumentBuilder builder, string format)
        {
            FieldTime field = (FieldTime)builder.InsertField(FieldType.FieldTime, true);
            builder.MoveTo(field.Separator);
            builder.Write(format);
            builder.MoveTo(field.Start.ParentNode);

            builder.InsertParagraph();
            return field;
        }
        //ExEnd

        private void TestFieldTime(Document doc)
        {
            DateTime docLoadingTime = DateTime.Now;
            doc = DocumentHelper.SaveOpen(doc);

            FieldTime field = (FieldTime)doc.Range.Fields[0];

            Assert.AreEqual(" TIME ", field.GetFieldCode());
            Assert.AreEqual(FieldType.FieldTime, field.Type);
            Assert.AreEqual(DateTime.Parse(field.Result), DateTime.Today.AddHours(docLoadingTime.Hour).AddMinutes(docLoadingTime.Minute));

            field = (FieldTime)doc.Range.Fields[1];

            Assert.AreEqual(" TIME \\@ HHmm", field.GetFieldCode());
            Assert.AreEqual(FieldType.FieldTime, field.Type);
            Assert.AreEqual(DateTime.Parse(field.Result), DateTime.Today.AddHours(docLoadingTime.Hour).AddMinutes(docLoadingTime.Minute));

            field = (FieldTime)doc.Range.Fields[2];

            Assert.AreEqual(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.GetFieldCode());
            Assert.AreEqual(FieldType.FieldTime, field.Type);
            Assert.AreEqual(DateTime.Parse(field.Result), DateTime.Today.AddHours(docLoadingTime.Hour).AddMinutes(docLoadingTime.Minute));
        }

        [Test]
        public void BidiOutline()
        {
            //ExStart
            //ExFor:FieldBidiOutline
            //ExFor:FieldShape
            //ExFor:FieldShape.Text
            //ExFor:ParagraphFormat.Bidi
            //ExSummary:Shows how to create RTL lists with BIDIOUTLINE fields.
            // Create a blank document and a document builder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use our builder to insert a BIDIOUTLINE field
            // This field numbers paragraphs like the AUTONUM/LISTNUM fields,
            // but is only visible when a RTL editing language is enabled, such as Hebrew or Arabic
            // The following field will display ".1", the RTL equivalent of list number "1."
            FieldBidiOutline field = (FieldBidiOutline)builder.InsertField(FieldType.FieldBidiOutline, true);
            builder.Writeln("שלום");

            Assert.AreEqual(" BIDIOUTLINE ", field.GetFieldCode());

            // Add two more BIDIOUTLINE fields, which will be automatically numbered ".2" and ".3"
            builder.InsertField(FieldType.FieldBidiOutline, true);
            builder.Writeln("שלום");
            builder.InsertField(FieldType.FieldBidiOutline, true);
            builder.Writeln("שלום");

            // Set the horizontal text alignment for every paragraph in the document to RTL
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.Bidi = true;
            }

            // If a RTL editing language is enabled in Microsoft Word, our fields will display numbers
            // Otherwise, they will appear as "###" 
            doc.Save(ArtifactsDir + "Field.BIDIOUTLINE.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Field.BIDIOUTLINE.docx");

            foreach (Field fieldBidiOutline in doc.Range.Fields)
                TestUtil.VerifyField(FieldType.FieldBidiOutline, " BIDIOUTLINE ", string.Empty, fieldBidiOutline);
        }

        [Test]
        public void Legacy()
        {
            //ExStart
            //ExFor:FieldEmbed
            //ExFor:FieldShape
            //ExFor:FieldShape.Text
            //ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled.
            // Open a document that was created in Microsoft Word 2003
            Document doc = new Document(MyDir + "Legacy fields.doc");

            // If we open the document in Word and press Alt+F9, we will see a SHAPE and an EMBED field
            // A SHAPE field is the anchor/canvas for an autoshape object with the "In line with text" wrapping style enabled
            // An EMBED field has the same function, but for an embedded object, such as a spreadsheet from an external Excel document
            // However, these fields will not appear in the document's Fields collection
            Assert.AreEqual(0, doc.Range.Fields.Count);

            // These fields are supported only by old versions of Microsoft Word
            // As such, they are converted into shapes during the document importation process and can instead be found in the collection of Shape nodes
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            Assert.AreEqual(3, shapes.Count);

            // The first Shape node corresponds to what was the SHAPE field in the input document: the inline canvas for an autoshape
            Shape shape = (Shape)shapes[0];
            Assert.AreEqual(ShapeType.Image, shape.ShapeType);

            // The next Shape node is the autoshape that is within the canvas
            shape = (Shape)shapes[1];
            Assert.AreEqual(ShapeType.Can, shape.ShapeType);

            // The third Shape is what was the EMBED field that contained the external spreadsheet
            shape = (Shape)shapes[2];
            Assert.AreEqual(ShapeType.OleObject, shape.ShapeType);
            //ExEnd
        }
    }
}