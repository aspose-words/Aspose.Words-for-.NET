// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.Data.OleDb;
using System.Net;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;
using Table = Aspose.Words.Tables.Table;
using Image = System.Drawing.Image;

namespace ApiExamples
{
    class TestUtil
    {
        /// <summary>
        /// Checks whether a field has an expected FieldType, field code and output result equal to their expected values.
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

#if NETFRAMEWORK || JAVA
        /// <summary>
        /// Checks whether an SQL query performed on a database file stored in the local file system
        /// produces a result that resembles an input Aspose.Words Table.
        /// </summary>
        /// <param name="expectedResult">Expected result of the SQL query in the form of an Aspose.Words table.</param>
        /// <param name="dbFilename">Local system filename of a database file.</param>
        /// <param name="sqlQuery">Microsoft.Jet.OLEDB.4.0-compliant SQL query.</param>
        internal static void TableMatchesQueryResult(Table expectedResult, string dbFilename, string sqlQuery)
        {
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
        }
#endif

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
        /// Checks whether a filename points to a valid image with specified dimensions.
        /// </summary>
        /// <remarks>
        /// Serves as a way to check that an image file is valid and nonempty without looking up its file size.
        /// </remarks>
        /// <param name="expectedWidth"></param>
        /// <param name="expectedHeight"></param>
        /// <param name="filename"></param>
        internal static void VerifyImage(int expectedWidth, int expectedHeight, string filename)
        {
            try
            {
                using (Image image = Image.FromFile(filename))
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
    }
}
