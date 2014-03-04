//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

using System;
using System.Text;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.IO;
using System.Data.OleDb;
using System.Reflection;


namespace ImportTableFromDataTableExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");
            // This is the location to our database. You must have the Examples folder extracted as well for the database to be found.
            string databaseDir = new Uri(new Uri(exeDir), "../../../../Examples/Common/Database/").LocalPath;

            // Create the output directory if it doesn't exist.
            if (!Directory.Exists(dataDir))
                Directory.CreateDirectory(dataDir);

            //ExStart
            //ExFor:Table.StyleIdentifier
            //ExFor:StyleIdentifier
            //ExFor:Table.StyleOptions
            //ExFor:TableStyleOptions
            //ExId:ImportDataTableCaller
            //ExSummary:Shows how to import the data from a DataTable and insert it into a new table in the document.
            // Create a new document.
            Document doc = new Document();

            // We can position where we want the table to be inserted and also specify any extra formatting to be
            // applied onto the table as well.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We want to rotate the page landscape as we expect a wide table.
            doc.FirstSection.PageSetup.Orientation = Orientation.Landscape;

            // Retrieve the data from our data source which is stored as a DataTable.
            DataTable dataTable = GetEmployees(databaseDir);

            // Build a table in the document from the data contained in the DataTable.
            Table table = ImportTableFromDataTable(builder, dataTable, true);

            // We can apply a table style as a very quick way to apply formatting to the entire table.
            table.StyleIdentifier = StyleIdentifier.MediumList2Accent1;
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands | TableStyleOptions.LastColumn;

            // For our table we want to remove the heading for the image column.
            table.FirstRow.LastCell.RemoveAllChildren();

            doc.Save(dataDir + "Table.FromDataTable Out.docx");
            //ExEnd
            
            // Do some verification on the generated table.
            doc.ExpandTableStylesToDirectFormatting();
            Debug.Assert(table.Rows.Count == 6, "Unexpected row count");
            Debug.Assert(doc.GetChildNodes(NodeType.Table, true).Count == 1, "Unexpected table count");
            Debug.Assert(table.FirstRow.FirstCell.ToString(SaveFormat.Text).Trim() == "EmployeeID", "Unexpected header text");
            Debug.Assert(table.Rows[2].Cells[2].ToString(SaveFormat.Text).Trim() == "Andrew", "Unexpected row text");
            Debug.Assert(table.Rows[1].FirstCell.CellFormat.Shading.BackgroundPatternColor != Color.Empty, "Unexpected cell shading");
        }

        //ExStart
        //ExId:ImportTableFromDataTable
        //ExSummary:Provides a method to import data from the DataTable and insert it into a new table using the DocumentBuilder.
        /// <summary>
        /// Imports the content from the specified DataTable into a new Aspose.Words Table object. 
        /// The table is inserted at the current position of the document builder and using the current builder's formatting if any is defined.
        /// </summary>
        public static Table ImportTableFromDataTable(DocumentBuilder builder, DataTable dataTable, bool importColumnHeadings)
        {
            Table table = builder.StartTable();

            // Check if the names of the columns from the data source are to be included in a header row.
            if (importColumnHeadings)
            {
                // Store the original values of these properties before changing them.
                bool boldValue = builder.Font.Bold;
                ParagraphAlignment paragraphAlignmentValue = builder.ParagraphFormat.Alignment;

                // Format the heading row with the appropriate properties.
                builder.Font.Bold = true;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                // Create a new row and insert the name of each column into the first row of the table.
                foreach (DataColumn column in dataTable.Columns)
                {
                    builder.InsertCell();
                    builder.Writeln(column.ColumnName);
                }

                builder.EndRow();

                // Restore the original formatting.
                builder.Font.Bold = boldValue;
                builder.ParagraphFormat.Alignment = paragraphAlignmentValue;
            }

            foreach (DataRow dataRow in dataTable.Rows)
            {
                foreach (object item in dataRow.ItemArray)
                {
                    // Insert a new cell for each object.
                    builder.InsertCell();

                    switch (item.GetType().Name)
                    {
                        case "Byte[]":
                            // Assume a byte array is an image. Other data types can be added here.
                            builder.InsertImage(GetImageFromByteArray((byte[])item), 50, 50);
                            break;
                        case "DateTime":
                            // Define a custom format for dates and times.
                            DateTime dateTime = (DateTime)item;
                            builder.Write(dateTime.ToString("MMMM d, yyyy"));
                            break;
                        default:
                            // By default any other item will be inserted as text.
                            builder.Write(item.ToString());
                            break;
                    }

                }

                // After we insert all the data from the current record we can end the table row.
                builder.EndRow();
            }

            // We have finished inserting all the data from the DataTable, we can end the table.
            builder.EndTable();

            return table;
        }
        //ExEnd

        /// <summary>
        /// Returns a .NET Image object from the specified byte array.
        /// </summary>
        private static Image GetImageFromByteArray(byte[] imageBytes)
        {
            // Some drivers can pick up some junk data to the start of binary storage fields.
            // This means we cannot directly read the bytes into an image, we first need
            // to skip past until we find the start of the image.
            string imageString = Encoding.ASCII.GetString(imageBytes);
            int index = imageString.IndexOf("BM");
            return Image.FromStream(new MemoryStream(imageBytes, index, imageBytes.Length - index));
        }

        /// <summary>
        /// Retrieves employee data from an external database.
        /// </summary>
        private static DataTable GetEmployees(string databaseDir)
        {
            // Open a database connection.
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                databaseDir + "Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            // Create the command.
            OleDbCommand cmd = new OleDbCommand("SELECT TOP 5 EmployeeID, LastName, FirstName, Title, Birthdate, Address, City, PhotoBLOB FROM Employees", conn);

            // Fill an ADO.NET table from the command.
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable table = new DataTable();
            da.Fill(table);

            // Close database.
            conn.Close();

            return table;
        }
    }
}