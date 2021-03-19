using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Tables
{
    internal class WorkingWithTables : DocsExamplesBase
    {
        [Test]
        public void RemoveColumn()
        {
            //ExStart:RemoveColumn
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 1, true);

            Column column = Column.FromIndex(table, 2);
            column.Remove();
            //ExEnd:RemoveColumn
        }

        [Test]
        public void InsertBlankColumn()
        {
            //ExStart:InsertBlankColumn
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            //ExStart:GetPlainText
            Column column = Column.FromIndex(table, 0);
            // Print the plain text of the column to the screen.
            Console.WriteLine(column.ToTxt());
            //ExEnd:GetPlainText
            
            // Create a new column to the left of this column.
            // This is the same as using the "Insert Column Before" command in Microsoft Word.
            Column newColumn = column.InsertColumnBefore();

            foreach (Cell cell in newColumn.Cells)
                cell.FirstParagraph.AppendChild(new Run(doc, "Column Text " + newColumn.IndexOf(cell)));
            //ExEnd:InsertBlankColumn
        }

        //ExStart:ColumnClass
        /// <summary>
        /// Represents a facade object for a column of a table in a Microsoft Word document.
        /// </summary>
        internal class Column
        {
            private Column(Table table, int columnIndex)
            {
                mTable = table ?? throw new ArgumentException("table");
                mColumnIndex = columnIndex;
            }

            /// <summary>
            /// Returns a new column facade from the table and supplied zero-based index.
            /// </summary>
            public static Column FromIndex(Table table, int columnIndex)
            {
                return new Column(table, columnIndex);
            }

            /// <summary>
            /// Returns the cells which make up the column.
            /// </summary>
            public Cell[] Cells => GetColumnCells().ToArray();

            /// <summary>
            /// Returns the index of the given cell in the column.
            /// </summary>
            public int IndexOf(Cell cell)
            {
                return GetColumnCells().IndexOf(cell);
            }

            /// <summary>
            /// Inserts a brand new column before this column into the table.
            /// </summary>
            public Column InsertColumnBefore()
            {
                Cell[] columnCells = Cells;

                if (columnCells.Length == 0)
                    throw new ArgumentException("Column must not be empty");

                // Create a clone of this column.
                foreach (Cell cell in columnCells)
                    cell.ParentRow.InsertBefore(cell.Clone(false), cell);

                // This is the new column.
                Column column = new Column(columnCells[0].ParentRow.ParentTable, mColumnIndex);

                // We want to make sure that the cells are all valid to work with (have at least one paragraph).
                foreach (Cell cell in column.Cells)
                    cell.EnsureMinimum();

                // Increase the index which this column represents since there is now one extra column in front.
                mColumnIndex++;

                return column;
            }

            /// <summary>
            /// Removes the column from the table.
            /// </summary>
            public void Remove()
            {
                foreach (Cell cell in Cells)
                    cell.Remove();
            }

            /// <summary>
            /// Returns the text of the column. 
            /// </summary>
            public string ToTxt()
            {
                StringBuilder builder = new StringBuilder();

                foreach (Cell cell in Cells)
                    builder.Append(cell.ToString(SaveFormat.Text));

                return builder.ToString();
            }

            /// <summary>
            /// Provides an up-to-date collection of cells which make up the column represented by this facade.
            /// </summary>
            private List<Cell> GetColumnCells()
            {
                List<Cell> columnCells = new List<Cell>();

                foreach (Row row in mTable.Rows)
                {
                    Cell cell = row.Cells[mColumnIndex];
                    if (cell != null)
                        columnCells.Add(cell);
                }

                return columnCells;
            }

            private int mColumnIndex;
            private readonly Table mTable;
        }
        //ExEnd:ColumnClass

        [Test]
        public void AutoFitTableToContents()
        {
            //ExStart:AutoFitTableToContents
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            doc.Save(ArtifactsDir + "WorkingWithTables.AutoFitTableToContents.docx");
            //ExEnd:AutoFitTableToContents
        }

        [Test]
        public void AutoFitTableToFixedColumnWidths()
        {
            //ExStart:AutoFitTableToFixedColumnWidths
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Disable autofitting on this table.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            doc.Save(ArtifactsDir + "WorkingWithTables.AutoFitTableToFixedColumnWidths.docx");
            //ExEnd:AutoFitTableToFixedColumnWidths
        }

        [Test]
        public void AutoFitTableToPageWidth()
        {
            //ExStart:AutoFitTableToPageWidth
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // Autofit the first table to the page width.
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            doc.Save(ArtifactsDir + "WorkingWithTables.AutoFitTableToWindow.docx");
            //ExEnd:AutoFitTableToPageWidth
        }

        [Test]
        public void BuildTableFromDataTable()
        {
            //ExStart:BuildTableFromDataTable
            Document doc = new Document();
            // We can position where we want the table to be inserted and specify any extra formatting to the table.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We want to rotate the page landscape as we expect a wide table.
            doc.FirstSection.PageSetup.Orientation = Orientation.Landscape;

            DataSet ds = new DataSet();
            ds.ReadXml(MyDir + "List of people.xml");
            // Retrieve the data from our data source, which is stored as a DataTable.
            DataTable dataTable = ds.Tables[0];

            // Build a table in the document from the data contained in the DataTable.
            Table table = ImportTableFromDataTable(builder, dataTable, true);

            // We can apply a table style as a very quick way to apply formatting to the entire table.
            table.StyleIdentifier = StyleIdentifier.MediumList2Accent1;
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands | TableStyleOptions.LastColumn;

            // For our table, we want to remove the heading for the image column.
            table.FirstRow.LastCell.RemoveAllChildren();

            doc.Save(ArtifactsDir + "WorkingWithTables.BuildTableFromDataTable.docx");
            //ExEnd:BuildTableFromDataTable
        }

        //ExStart:ImportTableFromDataTable
        /// <summary>
        /// Imports the content from the specified DataTable into a new Aspose.Words Table object.
        /// The table is inserted at the document builder's current position and using the current builder's formatting if any is defined.
        /// </summary>
        public Table ImportTableFromDataTable(DocumentBuilder builder, DataTable dataTable,
            bool importColumnHeadings)
        {
            Table table = builder.StartTable();

            // Check if the columns' names from the data source are to be included in a header row.
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
                            DateTime dateTime = (DateTime) item;
                            builder.Write(dateTime.ToString("MMMM d, yyyy"));
                            break;
                        default:
                            // By default any other item will be inserted as text.
                            builder.Write(item.ToString());
                            break;
                    }
                }

                // After we insert all the data from the current record, we can end the table row.
                builder.EndRow();
            }

            // We have finished inserting all the data from the DataTable, we can end the table.
            builder.EndTable();

            return table;
        }
        //ExEnd:ImportTableFromDataTable

        [Test]
        public void CloneCompleteTable()
        {
            //ExStart:CloneCompleteTable
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Clone the table and insert it into the document after the original.
            Table tableClone = (Table) table.Clone(true);
            table.ParentNode.InsertAfter(tableClone, table);

            // Insert an empty paragraph between the two tables,
            // or else they will be combined into one upon saving this has to do with document validation.
            table.ParentNode.InsertAfter(new Paragraph(doc), table);
            
            doc.Save(ArtifactsDir + "WorkingWithTables.CloneCompleteTable.docx");
            //ExEnd:CloneCompleteTable
        }

        [Test]
        public void CloneLastRow()
        {
            //ExStart:CloneLastRow
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            Row clonedRow = (Row) table.LastRow.Clone(true);
            // Remove all content from the cloned row's cells. This makes the row ready for new content to be inserted into.
            foreach (Cell cell in clonedRow.Cells)
                cell.RemoveAllChildren();

            table.AppendChild(clonedRow);

            doc.Save(ArtifactsDir + "WorkingWithTables.CloneLastRow.docx");
            //ExEnd:CloneLastRow
        }
        
        [Test]
        public void FindingIndex()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            //ExStart:RetrieveTableIndex
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
            int tableIndex = allTables.IndexOf(table);
            //ExEnd:RetrieveTableIndex
            Console.WriteLine("\nTable index is " + tableIndex);

            //ExStart:RetrieveRowIndex
            int rowIndex = table.IndexOf(table.LastRow);
            //ExEnd:RetrieveRowIndex
            Console.WriteLine("\nRow index is " + rowIndex);

            Row row = table.LastRow;
            //ExStart:RetrieveCellIndex
            int cellIndex = row.IndexOf(row.Cells[4]);
            //ExEnd:RetrieveCellIndex
            Console.WriteLine("\nCell index is " + cellIndex);
        }

        [Test]
        public void InsertTableDirectly()
        {
            //ExStart:InsertTableDirectly
            Document doc = new Document();
            
            // We start by creating the table object. Note that we must pass the document object
            // to the constructor of each node. This is because every node we create must belong
            // to some document.
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
            // to ensure that the specified node is valid. In this case, a valid table should have at least one Row and one cell.

            // Instead, we will handle creating the row and table ourselves.
            // This would be the best way to do this if we were creating a table inside an algorithm.
            Row row = new Row(doc);
            row.RowFormat.AllowBreakAcrossPages = true;
            table.AppendChild(row);

            // We can now apply any auto fit settings.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            Cell cell = new Cell(doc);
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            cell.CellFormat.Width = 80;
            cell.AppendChild(new Paragraph(doc));
            cell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 1 Text"));

            row.AppendChild(cell);

            // We would then repeat the process for the other cells and rows in the table.
            // We can also speed things up by cloning existing cells and rows.
            row.AppendChild(cell.Clone(false));
            row.LastCell.AppendChild(new Paragraph(doc));
            row.LastCell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 2 Text"));
            
            doc.Save(ArtifactsDir + "WorkingWithTables.InsertTableDirectly.docx");
            //ExEnd:InsertTableDirectly
        }

        [Test]
        public void InsertTableFromHtml()
        {
            //ExStart:InsertTableFromHtml
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Note that AutoFitSettings does not apply to tables inserted from HTML.
            builder.InsertHtml("<table>" +
                               "<tr>" +
                               "<td>Row 1, Cell 1</td>" +
                               "<td>Row 1, Cell 2</td>" +
                               "</tr>" +
                               "<tr>" +
                               "<td>Row 2, Cell 2</td>" +
                               "<td>Row 2, Cell 2</td>" +
                               "</tr>" +
                               "</table>");

            doc.Save(ArtifactsDir + "WorkingWithTables.InsertTableFromHtml.docx");
            //ExEnd:InsertTableFromHtml
        }

        [Test]
        public void CreateSimpleTable()
        {
            //ExStart:CreateSimpleTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Start building the table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content.");
            
            // Build the second cell.
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content.");
            
            // Call the following method to end the row and start a new row.
            builder.EndRow();

            // Build the first cell of the second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content");

            // Build the second cell.
            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content.");
            builder.EndRow();

            // Signal that we have finished building the table.
            builder.EndTable();

            doc.Save(ArtifactsDir + "WorkingWithTables.CreateSimpleTable.docx");
            //ExEnd:CreateSimpleTable
        }

        [Test]
        public void FormattedTable()
        {
            //ExStart:FormattedTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Table wide formatting must be applied after at least one row is present in the table.
            table.LeftIndent = 20.0;

            // Set height and define the height rule for the header row.
            builder.RowFormat.Height = 40.0;
            builder.RowFormat.HeightRule = HeightRule.AtLeast;

            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Size = 16;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;

            builder.CellFormat.Width = 100.0;
            builder.Write("Header Row,\n Cell 1");

            // We don't need to specify this cell's width because it's inherited from the previous cell.
            builder.InsertCell();
            builder.Write("Header Row,\n Cell 2");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Header Row,\n Cell 3");
            builder.EndRow();

            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            builder.CellFormat.Width = 100.0;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

            // Reset height and define a different height rule for table body.
            builder.RowFormat.Height = 30.0;
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.InsertCell();
            
            // Reset font formatting.
            builder.Font.Size = 12;
            builder.Font.Bold = false;

            builder.Write("Row 1, Cell 1 Content");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Row 1, Cell 3 Content");
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.Width = 100.0;
            builder.Write("Row 2, Cell 1 Content");

            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Row 2, Cell 3 Content.");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "WorkingWithTables.FormattedTable.docx");
            //ExEnd:FormattedTable
        }

        [Test]
        public void NestedTable()
        {
            //ExStart:NestedTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Cell cell = builder.InsertCell();
            builder.Writeln("Outer Table Cell 1");

            builder.InsertCell();
            builder.Writeln("Outer Table Cell 2");

            // This call is important to create a nested table within the first table. 
            // Without this call, the cells inserted below will be appended to the outer table.
            builder.EndTable();

            // Move to the first cell of the outer table.
            builder.MoveTo(cell.FirstParagraph);

            // Build the inner table.
            builder.InsertCell();
            builder.Writeln("Inner Table Cell 1");
            builder.InsertCell();
            builder.Writeln("Inner Table Cell 2");
            builder.EndTable();

            doc.Save(ArtifactsDir + "WorkingWithTables.NestedTable.docx");
            //ExEnd:NestedTable
        }

        [Test]
        public void CombineRows()
        {
            //ExStart:CombineRows
            Document doc = new Document(MyDir + "Tables.docx");

            // The rows from the second table will be appended to the end of the first table.
            Table firstTable = (Table) doc.GetChild(NodeType.Table, 0, true);
            Table secondTable = (Table) doc.GetChild(NodeType.Table, 1, true);

            // Append all rows from the current table to the next tables
            // with different cell count and widths can be joined into one table.
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            secondTable.Remove();

            doc.Save(ArtifactsDir + "WorkingWithTables.CombineRows.docx");
            //ExEnd:CombineRows
        }

        [Test]
        public void SplitTable()
        {
            //ExStart:SplitTable
            Document doc = new Document(MyDir + "Tables.docx");

            Table firstTable = (Table) doc.GetChild(NodeType.Table, 0, true);

            // We will split the table at the third row (inclusive).
            Row row = firstTable.Rows[2];

            // Create a new container for the split table.
            Table table = (Table) firstTable.Clone(false);

            // Insert the container after the original.
            firstTable.ParentNode.InsertAfter(table, firstTable);

            // Add a buffer paragraph to ensure the tables stay apart.
            firstTable.ParentNode.InsertAfter(new Paragraph(doc), firstTable);

            Row currentRow;

            do
            {
                currentRow = firstTable.LastRow;
                table.PrependChild(currentRow);
            } while (currentRow != row);

            doc.Save(ArtifactsDir + "WorkingWithTables.SplitTable.docx");
            //ExEnd:SplitTable
        }

        [Test]
        public void RowFormatDisableBreakAcrossPages()
        {
            //ExStart:RowFormatDisableBreakAcrossPages
            Document doc = new Document(MyDir + "Table spanning two pages.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Disable breaking across pages for all rows in the table.
            foreach (Row row in table.Rows)
                row.RowFormat.AllowBreakAcrossPages = false;

            doc.Save(ArtifactsDir + "WorkingWithTables.RowFormatDisableBreakAcrossPages.docx");
            //ExEnd:RowFormatDisableBreakAcrossPages
        }

        [Test]
        public void KeepTableTogether()
        {
            //ExStart:KeepTableTogether
            Document doc = new Document(MyDir + "Table spanning two pages.docx");
            
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // We need to enable KeepWithNext for every paragraph in the table to keep it from breaking across a page,
            // except for the last paragraphs in the last row of the table.
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                cell.EnsureMinimum();

                foreach (Paragraph para in cell.Paragraphs)
                    if (!(cell.ParentRow.IsLastRow && para.IsEndOfCell))
                        para.ParagraphFormat.KeepWithNext = true;
            }

            doc.Save(ArtifactsDir + "WorkingWithTables.KeepTableTogether.docx");
            //ExEnd:KeepTableTogether
        }

        [Test]
        public void CheckCellsMerged()
        {
            //ExStart:CheckCellsMerged 
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Console.WriteLine(PrintCellMergeType(cell));
                }
            }
            //ExEnd:CheckCellsMerged 
        }

        //ExStart:PrintCellMergeType 
        public string PrintCellMergeType(Cell cell)
        {
            bool isHorizontallyMerged = cell.CellFormat.HorizontalMerge != CellMerge.None;
            bool isVerticallyMerged = cell.CellFormat.VerticalMerge != CellMerge.None;
            
            string cellLocation =
                $"R{cell.ParentRow.ParentTable.IndexOf(cell.ParentRow) + 1}, C{cell.ParentRow.IndexOf(cell) + 1}";

            if (isHorizontallyMerged && isVerticallyMerged)
                return $"The cell at {cellLocation} is both horizontally and vertically merged";
            
            if (isHorizontallyMerged)
                return $"The cell at {cellLocation} is horizontally merged.";
            
            if (isVerticallyMerged)
                return $"The cell at {cellLocation} is vertically merged";
            
            return $"The cell at {cellLocation} is not merged";
        }
        //ExEnd:PrintCellMergeType
        
        [Test]
        public void VerticalMerge()
        {
            //ExStart:VerticalMerge           
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in one cell");
            builder.EndRow();

            builder.InsertCell();
            // This cell is vertically merged to the cell above and should be empty.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in another cell");
            builder.EndRow();
            builder.EndTable();
            
            doc.Save(ArtifactsDir + "WorkingWithTables.VerticalMerge.docx");
            //ExEnd:VerticalMerge
        }

        [Test]
        public void HorizontalMerge()
        {
            //ExStart:HorizontalMerge         
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            // This cell is merged to the previous and should be empty.
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Text in one cell.");

            builder.InsertCell();
            builder.Write("Text in another cell.");
            builder.EndRow();
            builder.EndTable();
            
            doc.Save(ArtifactsDir + "WorkingWithTables.HorizontalMerge.docx");
            //ExEnd:HorizontalMerge
        }

        [Test]
        public void MergeCellRange()
        {
            //ExStart:MergeCellRange
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            Table table = doc.FirstSection.Body.Tables[0];

            // We want to merge the range of cells found inbetween these two cells.
            Cell cellStartRange = table.Rows[0].Cells[0];
            Cell cellEndRange = table.Rows[1].Cells[1];

            // Merge all the cells between the two specified cells into one.
            MergeCells(cellStartRange, cellEndRange);
            
            doc.Save(ArtifactsDir + "WorkingWithTables.MergeCellRange.docx");
            //ExEnd:MergeCellRange
        }

        [Test]
        public void PrintHorizontalAndVerticalMerged()
        {
            //ExStart:PrintHorizontalAndVerticalMerged
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            SpanVisitor visitor = new SpanVisitor(doc);
            doc.Accept(visitor);
            //ExEnd:PrintHorizontalAndVerticalMerged
        }

        [Test]
        public void ConvertToHorizontallyMergedCells()
        {
            //ExStart:ConvertToHorizontallyMergedCells         
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            Table table = doc.FirstSection.Body.Tables[0];
            // Now merged cells have appropriate merge flags.
            table.ConvertToHorizontallyMergedCells();
            //ExEnd:ConvertToHorizontallyMergedCells
        }

        //ExStart:MergeCells
        internal void MergeCells(Cell startCell, Cell endCell)
        {
            Table parentTable = startCell.ParentRow.ParentTable;

            // Find the row and cell indices for the start and end cell.
            Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell),
                parentTable.IndexOf(startCell.ParentRow));
            Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow));

            // Create a range of cells to be merged based on these indices.
            // Inverse each index if the end cell is before the start cell.
            Rectangle mergeRange = new Rectangle(Math.Min(startCellPos.X, endCellPos.X),
                Math.Min(startCellPos.Y, endCellPos.Y),
                Math.Abs(endCellPos.X - startCellPos.X) + 1, Math.Abs(endCellPos.Y - startCellPos.Y) + 1);

            foreach (Row row in parentTable.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Point currentPos = new Point(row.IndexOf(cell), parentTable.IndexOf(row));

                    // Check if the current cell is inside our merge range, then merge it.
                    if (mergeRange.Contains(currentPos))
                    {
                        cell.CellFormat.HorizontalMerge = currentPos.X == mergeRange.X ? CellMerge.First : CellMerge.Previous;

                        cell.CellFormat.VerticalMerge = currentPos.Y == mergeRange.Y ? CellMerge.First : CellMerge.Previous;
                    }
                }
            }
        }
        //ExEnd:MergeCells
        
        //ExStart:HorizontalAndVerticalMergeHelperClasses
        /// <summary>
        /// Helper class that contains collection of rowinfo for each row.
        /// </summary>
        public class TableInfo
        {
            public List<RowInfo> Rows { get; } = new List<RowInfo>();
        }

        /// <summary>
        /// Helper class that contains collection of cellinfo for each cell.
        /// </summary>
        public class RowInfo
        {
            public List<CellInfo> Cells { get; } = new List<CellInfo>();
        }

        /// <summary>
        /// Helper class that contains info about cell. currently here is only colspan and rowspan.
        /// </summary>
        public class CellInfo
        {
            public CellInfo(int colSpan, int rowSpan)
            {
                ColSpan = colSpan;
                RowSpan = rowSpan;
            }

            public int ColSpan { get; }
            public int RowSpan { get; }
        }

        public class SpanVisitor : DocumentVisitor
        {
            /// <summary>
            /// Creates new SpanVisitor instance.
            /// </summary>
            /// <param name="doc">
            /// Is document which we should parse.
            /// </param>
            public SpanVisitor(Document doc)
            {
                mWordTables = doc.GetChildNodes(NodeType.Table, true);

                // We will parse HTML to determine the rowspan and colspan of each cell.
                MemoryStream htmlStream = new MemoryStream();

                Aspose.Words.Saving.HtmlSaveOptions options = new Aspose.Words.Saving.HtmlSaveOptions
                {
                    ImagesFolder = Path.GetTempPath()
                };

                doc.Save(htmlStream, options);

                // Load HTML into the XML document.
                XmlDocument xmlDoc = new XmlDocument();
                htmlStream.Position = 0;
                xmlDoc.Load(htmlStream);

                // Get collection of tables in the HTML document.
                XmlNodeList tables = xmlDoc.DocumentElement.SelectNodes("// Table");

                foreach (XmlNode table in tables)
                {
                    TableInfo tableInf = new TableInfo();
                    // Get collection of rows in the table.
                    XmlNodeList rows = table.SelectNodes("tr");

                    foreach (XmlNode row in rows)
                    {
                        RowInfo rowInf = new RowInfo();
                        // Get collection of cells.
                        XmlNodeList cells = row.SelectNodes("td");

                        foreach (XmlNode cell in cells)
                        {
                            // Determine row span and colspan of the current cell.
                            XmlAttribute colSpanAttr = cell.Attributes["colspan"];
                            XmlAttribute rowSpanAttr = cell.Attributes["rowspan"];

                            int colSpan = colSpanAttr == null ? 0 : int.Parse(colSpanAttr.Value);
                            int rowSpan = rowSpanAttr == null ? 0 : int.Parse(rowSpanAttr.Value);

                            CellInfo cellInf = new CellInfo(colSpan, rowSpan);
                            rowInf.Cells.Add(cellInf);
                        }

                        tableInf.Rows.Add(rowInf);
                    }

                    mTables.Add(tableInf);
                }
            }

            public override VisitorAction VisitCellStart(Cell cell)
            {
                int tabIdx = mWordTables.IndexOf(cell.ParentRow.ParentTable);
                int rowIdx = cell.ParentRow.ParentTable.IndexOf(cell.ParentRow);
                int cellIdx = cell.ParentRow.IndexOf(cell);

                int colSpan = 0;
                int rowSpan = 0;
                if (tabIdx < mTables.Count &&
                    rowIdx < mTables[tabIdx].Rows.Count &&
                    cellIdx < mTables[tabIdx].Rows[rowIdx].Cells.Count)
                {
                    colSpan = mTables[tabIdx].Rows[rowIdx].Cells[cellIdx].ColSpan;
                    rowSpan = mTables[tabIdx].Rows[rowIdx].Cells[cellIdx].RowSpan;
                }

                Console.WriteLine("{0}.{1}.{2} colspan={3}\t rowspan={4}", tabIdx, rowIdx, cellIdx, colSpan, rowSpan);

                return VisitorAction.Continue;
            }

            private readonly List<TableInfo> mTables = new List<TableInfo>();
            private readonly NodeCollection mWordTables;
        }
        //ExEnd:HorizontalAndVerticalMergeHelperClasses

        [Test]
        public void RepeatRowsOnSubsequentPages()
        {
            //ExStart:RepeatRowsOnSubsequentPages
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.RowFormat.HeadingFormat = true;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.Width = 100;
            builder.InsertCell();
            builder.Writeln("Heading row 1");
            builder.EndRow();
            builder.InsertCell();
            builder.Writeln("Heading row 2");
            builder.EndRow();

            builder.CellFormat.Width = 50;
            builder.ParagraphFormat.ClearFormatting();

            for (int i = 0; i < 50; i++)
            {
                builder.InsertCell();
                builder.RowFormat.HeadingFormat = false;
                builder.Write("Column 1 Text");
                builder.InsertCell();
                builder.Write("Column 2 Text");
                builder.EndRow();
            }

            doc.Save(ArtifactsDir + "WorkingWithTables.RepeatRowsOnSubsequentPages.docx");
            //ExEnd:RepeatRowsOnSubsequentPages
        }

        [Test]
        public void AutoFitToPageWidth()
        {
            //ExStart:AutoFitToPageWidth
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with a width that takes up half the page width.
            Table table = builder.StartTable();

            builder.InsertCell();
            table.PreferredWidth = PreferredWidth.FromPercent(50);
            builder.Writeln("Cell #1");

            builder.InsertCell();
            builder.Writeln("Cell #2");

            builder.InsertCell();
            builder.Writeln("Cell #3");

            doc.Save(ArtifactsDir + "WorkingWithTables.AutoFitToPageWidth.docx");
            //ExEnd:AutoFitToPageWidth
        }

        [Test]
        public void PreferredWidthSettings()
        {
            //ExStart:PreferredWidthSettings
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table row made up of three cells which have different preferred widths.
            builder.StartTable();

            // Insert an absolute sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.Writeln("Cell at 40 points width");

            // Insert a relative (percent) sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Writeln("Cell at 20% width");

            // Insert a auto sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Writeln(
                "Cell automatically sized. The size of this cell is calculated from the table preferred width.");
            builder.Writeln("In this case the cell will fill up the rest of the available space.");

            doc.Save(ArtifactsDir + "WorkingWithTables.PreferredWidthSettings.docx");
            //ExEnd:PreferredWidthSettings
        }

        [Test]
        public void RetrievePreferredWidthType()
        {
            //ExStart:RetrievePreferredWidthType
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            //ExStart:AllowAutoFit
            table.AllowAutoFit = true;
            //ExEnd:AllowAutoFit

            Cell firstCell = table.FirstRow.FirstCell;
            PreferredWidthType type = firstCell.CellFormat.PreferredWidth.Type;
            double value = firstCell.CellFormat.PreferredWidth.Value;
            //ExEnd:RetrievePreferredWidthType
        }

        [Test]
        public void GetTablePosition()
        {
            //ExStart:GetTablePosition
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            if (table.TextWrapping == TextWrapping.Around)
            {
                Console.WriteLine(table.RelativeHorizontalAlignment);
                Console.WriteLine(table.RelativeVerticalAlignment);
            }
            else
            {
                Console.WriteLine(table.Alignment);
            }
            //ExEnd:GetTablePosition
        }

        [Test]
        public void GetFloatingTablePosition()
        {
            //ExStart:GetFloatingTablePosition
            Document doc = new Document(MyDir + "Table wrapped by text.docx");
            
            foreach (Table table in doc.FirstSection.Body.Tables)
            {
                // If the table is floating type, then print its positioning properties.
                if (table.TextWrapping == TextWrapping.Around)
                {
                    Console.WriteLine(table.HorizontalAnchor);
                    Console.WriteLine(table.VerticalAnchor);
                    Console.WriteLine(table.AbsoluteHorizontalDistance);
                    Console.WriteLine(table.AbsoluteVerticalDistance);
                    Console.WriteLine(table.AllowOverlap);
                    Console.WriteLine(table.AbsoluteHorizontalDistance);
                    Console.WriteLine(table.RelativeVerticalAlignment);
                    Console.WriteLine("..............................");
                }
            }
            //ExEnd:GetFloatingTablePosition
        }

        [Test]
        public void FloatingTablePosition()
        {
            //ExStart:FloatingTablePosition
            Document doc = new Document(MyDir + "Table wrapped by text.docx");

            Table table = doc.FirstSection.Body.Tables[0];
            table.AbsoluteHorizontalDistance = 10;
            table.RelativeVerticalAlignment = VerticalAlignment.Center;

            doc.Save(ArtifactsDir + "WorkingWithTables.FloatingTablePosition.docx");
            //ExEnd:FloatingTablePosition
        }

        [Test]
        public void SetRelativeHorizontalOrVerticalPosition()
        {
            //ExStart:SetRelativeHorizontalOrVerticalPosition
            Document doc = new Document(MyDir + "Table wrapped by text.docx");

            Table table = doc.FirstSection.Body.Tables[0];
            table.HorizontalAnchor = RelativeHorizontalPosition.Column;
            table.VerticalAnchor = RelativeVerticalPosition.Page;

            doc.Save(ArtifactsDir + "WorkingWithTables.SetFloatingTablePosition.docx");
            //ExEnd:SetRelativeHorizontalOrVerticalPosition
        }
    }
}