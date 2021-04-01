// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Lists;
using Aspose.Words.Notes;
using NUnit.Framework;
using Table = Aspose.Words.Tables.Table;
using Image =
#if NET462 || JAVA
System.Drawing.Image;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Drawing;
#elif NETCOREAPP2_1 || __MOBILE__
SkiaSharp.SKBitmap;
using SkiaSharp;
#endif
using Shape = Aspose.Words.Drawing.Shape;

namespace ApiExamples
{
    class TestUtil : ApiExampleBase
    {
        /// <summary>
        /// Checks whether a file at a specified filename contains a valid image with specified dimensions.
        /// </summary>
        /// <remarks>
        /// Serves to check that an image file is valid and nonempty without looking up its file size.
        /// </remarks>
        /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
        /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
        /// <param name="filename">Local file system filename of the image file.</param>
        internal static void VerifyImage(int expectedWidth, int expectedHeight, string filename)
        {
            using (FileStream fileStream = new FileStream(filename, FileMode.Open))
            {
                VerifyImage(expectedWidth, expectedHeight, fileStream);
            }
        }

        /// <summary>
        /// Checks whether a stream contains a valid image with specified dimensions.
        /// </summary>
        /// <remarks>
        /// Serves to check that an image file is valid and nonempty without looking up its file size.
        /// </remarks>
        /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
        /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
        /// <param name="imageStream">Stream that contains the image.</param>
        internal static void VerifyImage(int expectedWidth, int expectedHeight, Stream imageStream)
        {
#if NET462 || JAVA
            using (Image image = Image.FromStream(imageStream))
#elif NETCOREAPP2_1 || __MOBILE__
            using (Image image = Image.Decode(imageStream))
#endif
            {
#if NET462 || NETCOREAPP2_1 || JAVA
                Assert.Multiple(() =>
                {
                    Assert.AreEqual(expectedWidth, image.Width, 1);
                    Assert.AreEqual(expectedHeight, image.Height, 1);
                });
#elif __MOBILE__
                Assert.AreEqual(expectedWidth, image.Width);
                Assert.AreEqual(expectedHeight, image.Height);
#endif
            }
        }

        /// <summary>
        /// Checks whether an image from the local file system contains any transparency.
        /// </summary>
        /// <param name="filename">Local file system filename of the image file.</param>
        internal static void ImageContainsTransparency(string filename)
        {
#if NET462 || JAVA
            using (Bitmap bitmap = (Bitmap)Image.FromFile(filename))
                for (int x = 0; x < bitmap.Width; x++)
                    for (int y = 0; y < bitmap.Height; y++)
                        if (bitmap.GetPixel(x, y).A != 255) return;

            Assert.Fail($"The image from \"{filename}\" does not contain any transparency.");
#elif NETCOREAPP2_1 || __MOBILE__
            using (Image image = Image.Decode(filename))
                foreach (SKColor pixelColor in image.Pixels)
                    if (pixelColor.Alpha != 255) return;

            Assert.Fail($"The image from \"{filename}\" does not contain any transparency.");
#endif
        }

        /// <summary>
        /// Checks whether an HTTP request sent to the specified address produces an expected web response. 
        /// </summary>
        /// <remarks>
        /// Serves as a notification of any URLs used in code examples becoming unusable in the future.
        /// </remarks>
        /// <param name="expectedHttpStatusCode">Expected result status code of a request HTTP "HEAD" method performed on the web address.</param>
        /// <param name="webAddress">URL where the request will be sent.</param>
        internal static void VerifyWebResponseStatusCode(HttpStatusCode expectedHttpStatusCode, string webAddress)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(webAddress);
            request.Method = "HEAD";

            Assert.AreEqual(expectedHttpStatusCode, ((HttpWebResponse)request.GetResponse()).StatusCode);
        }

        /// <summary>
        /// Checks whether an SQL query performed on a database file stored in the local file system
        /// produces a result that resembles the contents of an Aspose.Words table.
        /// </summary>
        /// <param name="expectedResult">Expected result of the SQL query in the form of an Aspose.Words table.</param>
        /// <param name="dbFilename">Local system filename of a database file.</param>
        /// <param name="sqlQuery">Microsoft.Jet.OLEDB.4.0-compliant SQL query.</param>
        internal static void TableMatchesQueryResult(Table expectedResult, string dbFilename, string sqlQuery)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            using (OleDbConnection connection = new OleDbConnection())
            {
                connection.ConnectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbFilename};";
                connection.Open();

                OleDbCommand command = connection.CreateCommand();
                command.CommandText = sqlQuery;
                OleDbDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

                DataTable myDataTable = new DataTable();
                myDataTable.Load(reader);

                Assert.AreEqual(expectedResult.Rows.Count, myDataTable.Rows.Count);
                Assert.AreEqual(expectedResult.Rows[0].Cells.Count, myDataTable.Columns.Count);

                for (int i = 0; i < myDataTable.Rows.Count; i++)
                    for (int j = 0; j < myDataTable.Columns.Count; j++)
                        Assert.AreEqual(expectedResult.Rows[i].Cells[j].GetText().Replace(ControlChar.Cell, string.Empty),
                            myDataTable.Rows[i][j].ToString());
            }
#endif
        }

        /// <summary>
        /// Checks whether a document produced during a mail merge contains every element of every table produced by a list of consecutive SQL queries on a database.
        /// </summary>
        /// <remarks>
        /// Currently, database types that cannot be represented by a string or a decimal are not checked for in the document.
        /// </remarks>
        /// <param name="dbFilename">Full local file system filename of a .mdb database file.</param>
        /// <param name="sqlQueries">List of SQL queries performed on the database all of whose results we expect to find in the document.</param>
        /// <param name="doc">Document created during a mail merge.</param>
        /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
        internal static void MailMergeMatchesQueryResultMultiple(string dbFilename, string[] sqlQueries, Document doc, bool onePagePerRow)
        {
            foreach (string query in sqlQueries)
                MailMergeMatchesQueryResult(dbFilename, query, doc, onePagePerRow);
        }

        /// <summary>
        /// Checks whether a document produced during a mail merge contains every element of a table produced by an SQL query on a database.
        /// </summary>
        /// <remarks>
        /// Currently, database types that cannot be represented by a string or a decimal are not checked for in the document.
        /// </remarks>
        /// <param name="dbFilename">Full local file system filename of a .mdb database file.</param>
        /// <param name="sqlQuery">SQL query performed on the database all of whose results we expect to find in the document.</param>
        /// <param name="doc">Document created during a mail merge.</param>
        /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
        internal static void MailMergeMatchesQueryResult(string dbFilename, string sqlQuery, Document doc, bool onePagePerRow)
        {
#if NET462 || JAVA
            List<string[]> expectedStrings = new List<string[]>(); 
            string connectionString = @"Driver={Microsoft Access Driver (*.mdb)};Dbq=" + dbFilename;

            using (OdbcConnection connection = new OdbcConnection())
            {
                connection.ConnectionString = connectionString;
                connection.Open();

                OdbcCommand command = connection.CreateCommand();
                command.CommandText = sqlQuery;

                using (OdbcDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    while (reader.Read())
                    {
                        string[] row = new string[reader.FieldCount];

                        for (int i = 0; i < reader.FieldCount; i++)
                            switch (reader[i])
                            {
                                case decimal d:
                                    row[i] = d.ToString("G29");
                                    break;
                                case string s:
                                    row[i] = s.Trim().Replace("\n", string.Empty);
                                    break;
                                default:
                                    row[i] = string.Empty;
                                    break;
                            }

                        expectedStrings.Add(row);
                    }
                }
            }

            MailMergeMatchesArray(expectedStrings.ToArray(), doc, onePagePerRow);
#endif
        }

        /// <summary>
        /// Checks whether a document produced during a mail merge contains every element of every DataTable in a DataSet.
        /// </summary>
        /// <param name="expectedResult">DataSet containing DataTables which contain values that we expect the document to contain.</param>
        /// <param name="doc">Document created during a mail merge.</param>
        /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
        internal static void MailMergeMatchesDataSet(DataSet dataSet, Document doc, bool onePagePerRow)
        {
            foreach (DataTable table in dataSet.Tables)
                MailMergeMatchesDataTable(table, doc, onePagePerRow);
        }

        /// <summary>
        /// Checks whether a document produced during a mail merge contains every element of a DataTable.
        /// </summary>
        /// <param name="expectedResult">Values from the mail merge data source that we expect the document to contain.</param>
        /// <param name="doc">Document created during a mail merge.</param>
        /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
        internal static void MailMergeMatchesDataTable(DataTable expectedResult, Document doc, bool onePagePerRow)
        {
            string[][] expectedStrings = new string[expectedResult.Rows.Count][];

            for (int i = 0; i < expectedResult.Rows.Count; i++)
                expectedStrings[i] = Array.ConvertAll(expectedResult.Rows[i].ItemArray, x => x.ToString());
            
            MailMergeMatchesArray(expectedStrings, doc, onePagePerRow);
        }

        /// <summary>
        /// Checks whether a document produced during a mail merge contains every element of an array of arrays of strings.
        /// </summary>
        /// <remarks>
        /// Only suitable for rectangular arrays.
        /// </remarks>
        /// <param name="expectedResult">Values from the mail merge data source that we expect the document to contain.</param>
        /// <param name="doc">Document created during a mail merge.</param>
        /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
        internal static void MailMergeMatchesArray(string[][] expectedResult, Document doc, bool onePagePerRow)
        {
            try
            {
                if (onePagePerRow)
                {
                    string[] docTextByPages = doc.GetText().Trim().Split(new[] { ControlChar.PageBreak }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0; i < expectedResult.Length; i++)
                        for (int j = 0; j < expectedResult[0].Length; j++)
                            if (!docTextByPages[i].Contains(expectedResult[i][j])) throw new ArgumentException(expectedResult[i][j]);
                }
                else
                {
                    string docText = doc.GetText();

                    for (int i = 0; i < expectedResult.Length; i++)
                        for (int j = 0; j < expectedResult[0].Length; j++)
                            if (!docText.Contains(expectedResult[i][j])) throw new ArgumentException(expectedResult[i][j]);

                }
            }
            catch (ArgumentException e)
            {
                Assert.Fail($"String \"{e.Message}\" not found in {(doc.OriginalFileName == null ? "a virtual document" : doc.OriginalFileName.Split('\\').Last())}.");
            }
        }

        /// <summary>
        /// Checks whether a file inside a document's OOXML package contains a string.
        /// </summary>
        /// <remarks>
        /// If an output document does not have a testable value that can be found as a property in its object when loaded,
        /// the value can sometimes be found in the document's OOXML package. 
        /// </remarks>
        /// <param name="expected">The string we are looking for.</param>
        /// <param name="docFilename">Local file system filename of the document.</param>
        /// <param name="docPartFilename">Name of the file within the document opened as a .zip that is expected to contain the string.</param>
        internal static void DocPackageFileContainsString(string expected, string docFilename, string docPartFilename)
        {
            using (ZipArchive archive = ZipFile.Open(docFilename, ZipArchiveMode.Update))
            {
                ZipArchiveEntry entry = archive.Entries.First(e => e.Name == docPartFilename);
                
                using (Stream stream = entry.Open())
                {
                   StreamContainsString(expected, stream);
                }
            }
        }

        /// <summary>
        /// Checks whether a file in the local file system contains a string in its raw data.
        /// </summary>
        /// <param name="expected">The string we are looking for.</param>
        /// <param name="filename">Local system filename of a file which, when read from the beginning, should contain the string.</param>
        internal static void FileContainsString(string expected, string filename)
        {
            if (!IsRunningOnMono())
            {
                using (Stream stream = new FileStream(filename, FileMode.Open))
                {
                    StreamContainsString(expected, stream);
                }
            }
        }

        /// <summary>
        /// Checks whether a stream contains a string.
        /// </summary>
        /// <param name="expected">The string we are looking for.</param>
        /// <param name="stream">The stream which, when read from the beginning, should contain the string.</param>
        private static void StreamContainsString(string expected, Stream stream)
        {
            char[] expectedSequence = expected.ToCharArray();

            long sequenceMatchLength = 0;
            while (stream.Position < stream.Length)
            {
                if ((char)stream.ReadByte() == expectedSequence[sequenceMatchLength])
                    sequenceMatchLength++;
                else
                    sequenceMatchLength = 0;

                if (sequenceMatchLength >= expectedSequence.Length)
                {
                    return;
                }
            }

            Assert.Fail($"String \"{(expected.Length <= 100 ? expected : expected.Substring(0, 100) + "...")}\" not found in the provided source.");
        }

        /// <summary>
        /// Checks whether values of properties of a field with a type not related to date/time are equal to expected values.
        /// </summary>
        /// <remarks>
        /// Best used when there are many fields closely being tested and should be avoided if a field has a long field code/result.
        /// </remarks>
        /// <param name="expectedType">The FieldType that we expect the field to have.</param>
        /// <param name="expectedFieldCode">The expected output value of GetFieldCode() being called on the field.</param>
        /// <param name="expectedResult">The field's expected result, which will be the value displayed by it in the document.</param>
        /// <param name="field">The field that's being tested.</param>
        internal static void VerifyField(FieldType expectedType, string expectedFieldCode, string expectedResult, Field field)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.AreEqual(expectedType, field.Type);
                Assert.AreEqual(expectedFieldCode, field.GetFieldCode(true));
                Assert.AreEqual(expectedResult, field.Result);
            });
#elif __MOBILE__
            Assert.AreEqual(expectedType, field.Type);
            Assert.AreEqual(expectedFieldCode, field.GetFieldCode(true));
            Assert.AreEqual(expectedResult, field.Result);
#endif
        }

        /// <summary>
        /// Checks whether values of properties of a field with a type related to date/time are equal to expected values.
        /// </summary>
        /// <remarks>
        /// Used when comparing DateTime instances to Field.Result values parsed to DateTime, which may differ slightly. 
        /// Give a delta value that's generous enough for any lower end system to pass, also a delta of zero is allowed.
        /// </remarks>
        /// <param name="expectedType">The FieldType that we expect the field to have.</param>
        /// <param name="expectedFieldCode">The expected output value of GetFieldCode() being called on the field.</param>
        /// <param name="expectedResult">The date/time that the field's result is expected to represent.</param>
        /// <param name="field">The field that's being tested.</param>
        /// <param name="delta">Margin of error for expectedResult.</param>
        internal static void VerifyField(FieldType expectedType, string expectedFieldCode, DateTime expectedResult, Field field, TimeSpan delta)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.AreEqual(expectedType, field.Type);
                Assert.AreEqual(expectedFieldCode, field.GetFieldCode(true));
                Assert.True(DateTime.TryParse(field.Result, out DateTime actual));

                if (field.Type == FieldType.FieldTime)
                    VerifyDate(expectedResult, actual, delta);
                else
                    VerifyDate(expectedResult.Date, actual, delta);
            });
#elif __MOBILE__
            Assert.AreEqual(expectedType, field.Type);
            Assert.AreEqual(expectedFieldCode, field.GetFieldCode(true));
            Assert.True(DateTime.TryParse(field.Result, out DateTime actual));

            if (field.Type == FieldType.FieldTime)
                VerifyDate(expectedResult, actual, delta);
            else
                VerifyDate(expectedResult.Date, actual, delta);
#endif
        }

        /// <summary>
        /// Checks whether a DateTime matches an expected value, with a margin of error.
        /// </summary>
        /// <param name="expected">The date/time that we expect the result to be.</param>
        /// <param name="actual">The DateTime object being tested.</param>
        /// <param name="delta">Margin of error for expectedResult.</param>
        internal static void VerifyDate(DateTime expected, DateTime actual, TimeSpan delta)
        {
            Assert.True(expected - actual <= delta);
        }

        /// <summary>
        /// Checks whether a field contains another complete field as a sibling within its nodes.
        /// </summary>
        /// <remarks>
        /// If two fields have the same immediate parent node and therefore their nodes are siblings,
        /// the FieldStart of the outer field appears before the FieldStart of the inner node,
        /// and the FieldEnd of the outer node appears after the FieldEnd of the inner node,
        /// then the inner field is considered to be nested within the outer field. 
        /// </remarks>
        /// <param name="innerField">The field that we expect to be fully within outerField.</param>
        /// <param name="outerField">The field that we to contain innerField.</param>
        internal static void FieldsAreNested(Field innerField, Field outerField)
        {
            CompositeNode innerFieldParent = innerField.Start.ParentNode;

            Assert.True(innerFieldParent == outerField.Start.ParentNode);
            Assert.True(innerFieldParent.ChildNodes.IndexOf(innerField.Start) > innerFieldParent.ChildNodes.IndexOf(outerField.Start));
            Assert.True(innerFieldParent.ChildNodes.IndexOf(innerField.End) < innerFieldParent.ChildNodes.IndexOf(outerField.End));
        }

        /// <summary>
        /// Checks whether a shape contains a valid image with specified dimensions.
        /// </summary>
        /// <remarks>
        /// Serves to check that an image file is valid and nonempty without looking up its data length.
        /// </remarks>
        /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
        /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
        /// <param name="expectedImageType">Expected format of the image.</param>
        /// <param name="imageShape">Shape that contains the image.</param>
        internal static void VerifyImageInShape(int expectedWidth, int expectedHeight, ImageType expectedImageType, Shape imageShape)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.True(imageShape.HasImage);
                Assert.AreEqual(expectedImageType, imageShape.ImageData.ImageType);
                Assert.AreEqual(expectedWidth, imageShape.ImageData.ImageSize.WidthPixels);
                Assert.AreEqual(expectedHeight, imageShape.ImageData.ImageSize.HeightPixels);
            });
#elif __MOBILE__
            Assert.True(imageShape.HasImage);
            Assert.AreEqual(expectedImageType, imageShape.ImageData.ImageType);
            Assert.AreEqual(expectedWidth, imageShape.ImageData.ImageSize.WidthPixels);
            Assert.AreEqual(expectedHeight, imageShape.ImageData.ImageSize.HeightPixels);
#endif
        }

        /// <summary>
        /// Checks whether values of a footnote's properties are equal to their expected values.
        /// </summary>
        /// <param name="expectedFootnoteType">Expected type of the footnote/endnote.</param>
        /// <param name="expectedIsAuto">Expected auto-numbered status of this footnote.</param>
        /// <param name="expectedReferenceMark">If "IsAuto" is false, then the footnote is expected to display this string instead of a number after referenced text.</param>
        /// <param name="expectedContents">Expected side comment provided by the footnote.</param>
        /// <param name="footnote">Footnote node in question.</param>
        internal static void VerifyFootnote(FootnoteType expectedFootnoteType, bool expectedIsAuto, string expectedReferenceMark, string expectedContents, Footnote footnote)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.AreEqual(expectedFootnoteType, footnote.FootnoteType);
                Assert.AreEqual(expectedIsAuto, footnote.IsAuto);
                Assert.AreEqual(expectedReferenceMark, footnote.ReferenceMark);
                Assert.AreEqual(expectedContents, footnote.ToString(SaveFormat.Text).Trim());
            });
#elif __MOBILE__
            Assert.AreEqual(expectedFootnoteType, footnote.FootnoteType);
            Assert.AreEqual(expectedIsAuto, footnote.IsAuto);
            Assert.AreEqual(expectedReferenceMark, footnote.ReferenceMark);
            Assert.AreEqual(expectedContents, footnote.ToString(SaveFormat.Text).Trim());
#endif
        }

        /// <summary>
        /// Checks whether values of a list level's properties are equal to their expected values.
        /// </summary>
        /// <remarks>
        /// Only necessary for list levels that have been explicitly created by the user.
        /// </remarks>
        /// <param name="expectedListFormat">Expected format for the list symbol.</param>
        /// <param name="expectedNumberPosition">Expected indent for this level, usually growing larger with each level.</param>
        /// <param name="expectedNumberStyle"></param>
        /// <param name="listLevel">List level in question.</param>
        internal static void VerifyListLevel(string expectedListFormat, double expectedNumberPosition, NumberStyle expectedNumberStyle, ListLevel listLevel)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.AreEqual(expectedListFormat, listLevel.NumberFormat);
                Assert.AreEqual(expectedNumberPosition, listLevel.NumberPosition);
                Assert.AreEqual(expectedNumberStyle, listLevel.NumberStyle);
            });
#elif __MOBILE__
            Assert.AreEqual(expectedListFormat, listLevel.NumberFormat);
            Assert.AreEqual(expectedNumberPosition, listLevel.NumberPosition);
            Assert.AreEqual(expectedNumberStyle, listLevel.NumberStyle);
#endif
        }
        
        /// <summary>
        /// Copies from the current position in src stream till the end.
        /// Copies into the current position in dst stream.
        /// </summary>
        internal static void CopyStream(Stream srcStream, Stream dstStream)
        {
            if (srcStream == null)
                throw new ArgumentNullException("srcStream");
            if (dstStream == null)
                throw new ArgumentNullException("dstStream");

            byte[] buf = new byte[65536];
            while (true)
            {
                int bytesRead = srcStream.Read(buf, 0, buf.Length);
                // Read returns 0 when reached end of stream
                // Checking for negative too to make it conceptually close to Java
                if (bytesRead <= 0)
                    break;
                dstStream.Write(buf, 0, bytesRead);
            }
        }
        
        /// <summary>
        /// Dumps byte array into a string.
        /// </summary>
        public static string DumpArray(byte[] data, int start, int count)
        {
            if (data == null)
                return "Null";

            StringBuilder builder = new StringBuilder();
            while (count > 0)
            {
                builder.AppendFormat("{0:X2} ", data[start]);
                start++;
                count--;
            }
            return builder.ToString();
        }

        /// <summary>
        /// Checks whether values of a tab stop's properties are equal to their expected values.
        /// </summary>
        /// <param name="expectedPosition">Expected position on the tab stop ruler, in points.</param>
        /// <param name="expectedTabAlignment">Expected position where the position is measured from </param>
        /// <param name="expectedTabLeader">Expected characters that pad the space between the start and end of the tab whitespace.</param>
        /// <param name="isClear">Whether or no this tab stop clears any tab stops.</param>
        /// <param name="tabStop">Tab stop that's being tested.</param>
        internal static void VerifyTabStop(double expectedPosition, TabAlignment expectedTabAlignment, TabLeader expectedTabLeader, bool isClear, TabStop tabStop)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.AreEqual(expectedPosition, tabStop.Position);
                Assert.AreEqual(expectedTabAlignment, tabStop.Alignment);
                Assert.AreEqual(expectedTabLeader, tabStop.Leader);
                Assert.AreEqual(isClear, tabStop.IsClear);
            });
#elif __MOBILE__
            Assert.AreEqual(expectedPosition, tabStop.Position);
            Assert.AreEqual(expectedTabAlignment, tabStop.Alignment);
            Assert.AreEqual(expectedTabLeader, tabStop.Leader);
            Assert.AreEqual(isClear, tabStop.IsClear);
#endif
        }

        /// <summary>
        /// Checks whether values of a shape's properties are equal to their expected values.
        /// </summary>
        /// <remarks>
        /// All dimension measurements are in points.
        /// </remarks>
        internal static void VerifyShape(ShapeType expectedShapeType, string expectedName, double expectedWidth, double expectedHeight, double expectedTop, double expectedLeft, Shape shape)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.AreEqual(expectedShapeType, shape.ShapeType);
                Assert.AreEqual(expectedName, shape.Name);
                Assert.AreEqual(expectedWidth, shape.Width);
                Assert.AreEqual(expectedHeight, shape.Height);
                Assert.AreEqual(expectedTop, shape.Top);
                Assert.AreEqual(expectedLeft, shape.Left);
            });
#elif __MOBILE__
            Assert.AreEqual(expectedShapeType, shape.ShapeType);
            Assert.AreEqual(expectedName, shape.Name);
            Assert.AreEqual(expectedWidth, shape.Width);
            Assert.AreEqual(expectedHeight, shape.Height);
            Assert.AreEqual(expectedTop, shape.Top);
            Assert.AreEqual(expectedLeft, shape.Left);
#endif
        }

        /// <summary>
        /// Checks whether values of properties of a textbox are equal to their expected values.
        /// </summary>
        /// <remarks>
        /// All dimension measurements are in points.
        /// </remarks>
        internal static void VerifyTextBox(LayoutFlow expectedLayoutFlow, bool expectedFitShapeToText, TextBoxWrapMode expectedTextBoxWrapMode, double marginTop, double marginBottom, double marginLeft, double marginRight, TextBox textBox)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.AreEqual(expectedLayoutFlow, textBox.LayoutFlow);
                Assert.AreEqual(expectedFitShapeToText, textBox.FitShapeToText);
                Assert.AreEqual(expectedTextBoxWrapMode, textBox.TextBoxWrapMode);
                Assert.AreEqual(marginTop, textBox.InternalMarginTop);
                Assert.AreEqual(marginBottom, textBox.InternalMarginBottom);
                Assert.AreEqual(marginLeft, textBox.InternalMarginLeft);
                Assert.AreEqual(marginRight, textBox.InternalMarginRight);
            });
#elif __MOBILE__
            Assert.AreEqual(expectedLayoutFlow, textBox.LayoutFlow);
            Assert.AreEqual(expectedFitShapeToText, textBox.FitShapeToText);
            Assert.AreEqual(expectedTextBoxWrapMode, textBox.TextBoxWrapMode);
            Assert.AreEqual(marginTop, textBox.InternalMarginTop);
            Assert.AreEqual(marginBottom, textBox.InternalMarginBottom);
            Assert.AreEqual(marginLeft, textBox.InternalMarginLeft);
            Assert.AreEqual(marginRight, textBox.InternalMarginRight);
#endif
        }

        /// <summary>
        /// Checks whether values of properties of an editable range are equal to their expected values.
        /// </summary>
        internal static void VerifyEditableRange(int expectedId, string expectedEditorUser, EditorType expectedEditorGroup, EditableRange editableRange)
        {
#if NET462 || NETCOREAPP2_1 || JAVA
            Assert.Multiple(() =>
            {
                Assert.AreEqual(expectedId, editableRange.Id);
                Assert.AreEqual(expectedEditorUser, editableRange.SingleUser);
                Assert.AreEqual(expectedEditorGroup, editableRange.EditorGroup);
            });
#elif __MOBILE__
            Assert.AreEqual(expectedId, editableRange.Id);
            Assert.AreEqual(expectedEditorUser, editableRange.SingleUser);
            Assert.AreEqual(expectedEditorGroup, editableRange.EditorGroup);
#endif
        }
    }
}
