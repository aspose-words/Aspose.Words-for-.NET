using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    public class BuildTableFromDataTable
    {
        public static void Run()
        {
            // ExStart:BuildTableFromDataTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();

            // Create a new document.
            Document doc = new Document();

            // We can position where we want the table to be inserted and also specify any extra formatting to be
            // applied onto the table as well.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We want to rotate the page landscape as we expect a wide table.
            doc.FirstSection.PageSetup.Orientation = Orientation.Landscape;

            DataSet ds = new DataSet();
            ds.ReadXml(dataDir + "Employees.xml");
            // Retrieve the data from our data source which is stored as a DataTable.
            DataTable dataTable = ds.Tables[0];

            // Build a table in the document from the data contained in the DataTable.
            Table table = ImportTableFromDataTable(builder, dataTable, true);

            // We can apply a table style as a very quick way to apply formatting to the entire table.
            table.StyleIdentifier = StyleIdentifier.MediumList2Accent1;
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands | TableStyleOptions.LastColumn;

            // For our table we want to remove the heading for the image column.
            table.FirstRow.LastCell.RemoveAllChildren();

            // Save the output document.
            doc.Save(dataDir + "Table.FromDataTable Out.docx");
            // ExEnd:BuildTableFromDataTable
            Console.WriteLine("\nDocument created successfully.\nFile saved at " + dataDir);
        }

        // ExStart:ImportTableFromDataTable
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
        // ExEnd:ImportTableFromDataTable
    }
}
