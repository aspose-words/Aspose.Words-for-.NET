// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Net;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Lists;
using NUnit.Framework;
using Table = Aspose.Words.Tables.Table;
using Image =
#if NET462 || JAVA
System.Drawing.Image;
using System.Collections.Generic;
using System.Data.Odbc;
#elif NETCOREAPP2_1 || __MOBILE__
SkiaSharp.SKBitmap;
using SkiaSharp;
#endif
using Shape = Aspose.Words.Drawing.Shape;

namespace ApiExamples
{
    class TestUtil
    {
        /// <summary>
        /// Checks whether a filename points to a valid image with specified dimensions.
        /// </summary>
        /// <remarks>
        /// Serves as a way to check that an image file is valid and nonempty without looking up its file size.
        /// </remarks>
        /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
        /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
        /// <param name="filename">Local file system filename of the image file.</param>
        internal static void VerifyImage(int expectedWidth, int expectedHeight, string filename)
        {
            try
            {
                #if NET462 || JAVA
                using (Image image = Image.FromFile(filename))
                #elif NETCOREAPP2_1 || __MOBILE__
                using (Image image = Image.Decode(filename))
                #endif
                {
                    Assert.AreEqual(expectedWidth, image.Width);
                    Assert.AreEqual(expectedHeight, image.Height);
                }
            }
            catch (OutOfMemoryException e)
            {
                Assert.Fail($"No valid image in this location:\n{filename}");
            }
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
            #if !__MOBILE__
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
        /// Checks whether values of a field's attributes are equal to expected values.
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
            Assert.AreEqual(expectedType, field.Type);
            Assert.AreEqual(expectedFieldCode, field.GetFieldCode(true));
            Assert.AreEqual(expectedResult, field.Result);
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
        /// Serves as a way to check that an image file is valid and nonempty without looking up its data length.
        /// </remarks>
        /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
        /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
        /// <param name="expectedImageType">Expected format of the image.</param>
        /// <param name="imageShape">Shape that contains the image.</param>
        internal static void VerifyImage(int expectedWidth, int expectedHeight, ImageType expectedImageType, Shape imageShape)
        {
            Assert.True(imageShape.HasImage);
            Assert.AreEqual(expectedImageType, imageShape.ImageData.ImageType);
            Assert.AreEqual(expectedWidth, imageShape.ImageData.ImageSize.WidthPixels);
            Assert.AreEqual(expectedHeight, imageShape.ImageData.ImageSize.HeightPixels);
        }

        /// <summary>
        /// Checks whether values of a footnote's attributes are equal to their expected values.
        /// </summary>
        /// <param name="expectedFootnoteType">Expected type of the footnote/endnote.</param>
        /// <param name="expectedIsAuto">Expected auto-numbered status of this footnote.</param>
        /// <param name="expectedReferenceMark">If "IsAuto" is false, then the footnote is expected to display this string instead of a number after referenced text.</param>
        /// <param name="expectedContents">Expected side comment provided by the footnote.</param>
        /// <param name="footnote">Footnote node in question.</param>
        internal static void VerifyFootnote(FootnoteType expectedFootnoteType, bool expectedIsAuto, string expectedReferenceMark, string expectedContents, Footnote footnote)
        {
            Assert.AreEqual(expectedFootnoteType, footnote.FootnoteType);
            Assert.AreEqual(expectedIsAuto, footnote.IsAuto);
            Assert.AreEqual(expectedReferenceMark, footnote.ReferenceMark);
            Assert.AreEqual(expectedContents, footnote.ToString(SaveFormat.Text).Trim());
        }

        /// <summary>
        /// Checks whether values of a list level's attributes are equal to their expected values.
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
            Assert.AreEqual(expectedListFormat, listLevel.NumberFormat);
            Assert.AreEqual(expectedNumberPosition, listLevel.NumberPosition);
            Assert.AreEqual(expectedNumberStyle, listLevel.NumberStyle);
        }
    }
}
