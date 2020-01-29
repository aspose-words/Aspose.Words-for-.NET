// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Examples using tables in documents.
    /// </summary>
    [TestFixture]
    public class ExTable : ApiExampleBase
    {
        [Test]
        public void DisplayContentOfTables()
        {
            //ExStart
            //ExFor:Cell
            //ExFor:CellCollection
            //ExFor:CellCollection.Item(System.Int32)
            //ExFor:CellCollection.ToArray
            //ExFor:Row
            //ExFor:Row.Cells
            //ExFor:RowCollection
            //ExFor:RowCollection.Item(System.Int32)
            //ExFor:RowCollection.ToArray
            //ExFor:Table
            //ExFor:Table.Rows
            //ExFor:TableCollection.Item(System.Int32)
            //ExFor:TableCollection.ToArray
            //ExSummary:Shows how to iterate through all tables in the document and display the content from each cell.
            Document doc = new Document(MyDir + "Tables.docx");

            // Here we get all tables from the Document node. You can do this for any other composite node
            // which can contain block level nodes. For example you can retrieve tables from header or from a cell
            // containing another table (nested tables)
            TableCollection tables = doc.FirstSection.Body.Tables;

            // We can make a new array to clone all of the tables in the collection
            Assert.AreEqual(2, tables.ToArray().Length);

            // Iterate through all tables in the document
            for (int i = 0; i < tables.Count; i++)
            {
                // Get the index of the table node as contained in the parent node of the table
                Console.WriteLine($"Start of Table {i}");

                RowCollection rows = tables[i].Rows;

                // RowCollections can be cloned into arrays
                Assert.AreEqual(rows, rows.ToArray());
                Assert.AreNotSame(rows, rows.ToArray());

                // Iterate through all rows in the table
                for (int j = 0; j < rows.Count; j++)
                {
                    Console.WriteLine($"\tStart of Row {j}");

                    CellCollection cells = rows[j].Cells;

                    // RowCollections can also be cloned into arrays 
                    Assert.AreEqual(cells, cells.ToArray());
                    Assert.AreNotSame(cells, cells.ToArray());

                    // Iterate through all cells in the row
                    for (int k = 0; k < cells.Count; k++)
                    {
                        // Get the plain text content of this cell
                        string cellText = cells[k].ToString(SaveFormat.Text).Trim();
                        // Print the content of the cell
                        Console.WriteLine($"\t\tContents of Cell:{k} = \"{cellText}\"");
                    }

                    Console.WriteLine($"\tEnd of Row {j}");
                }

                Console.WriteLine($"End of Table {i}\n");
            }
            //ExEnd

            Assert.That(tables.Count, Is.GreaterThan(0));
        }

        //ExStart
        //ExFor:Node.GetAncestor(NodeType)
        //ExFor:Node.GetAncestor(System.Type)
        //ExFor:Table.NodeType
        //ExFor:Cell.Tables
        //ExFor:TableCollection
        //ExFor:NodeCollection.Count
        //ExSummary:Shows how to find out if a table contains another table or if the table itself is nested inside another table.
        [Test] //ExSkip
        public void CalculateDepthOfNestedTables()
        {
            Document doc = new Document(MyDir + "Nested tables.docx");
            int tableIndex = 0;

            foreach (Table table in doc.GetChildNodes(NodeType.Table, true).OfType<Table>())
            {
                // First lets find if any cells in the table have tables themselves as children
                int count = GetChildTableCount(table);
                Console.WriteLine("Table #{0} has {1} tables directly within its cells", tableIndex, count);

                // Now let's try the other way around, lets try find if the table is nested inside another table and at what depth
                int tableDepth = GetNestedDepthOfTable(table);

                if (tableDepth > 0)
                    Console.WriteLine("Table #{0} is nested inside another table at depth of {1}", tableIndex,
                        tableDepth);
                else
                    Console.WriteLine("Table #{0} is a non nested table (is not a child of another table)", tableIndex);

                tableIndex++;
            }
        }

        /// <summary>
        /// Calculates what level a table is nested inside other tables.
        /// <returns>
        /// An integer containing the level the table is nested at.
        /// 0 = Table is not nested inside any other table
        /// 1 = Table is nested within one parent table
        /// 2 = Table is nested within two parent tables etc..</returns>
        /// </summary>
        private static int GetNestedDepthOfTable(Table table)
        {
            int depth = 0;

            // The parent of the table will be a Cell, instead attempt to find a grandparent that is of type Table
            Node parent = table.GetAncestor(table.NodeType);

            while (parent != null)
            {
                // Every time we find a table a level up we increase the depth counter and then try to find an
                // ancestor of type table from the parent
                depth++;
                parent = parent.GetAncestor(typeof(Table));
            }

            return depth;
        }

        /// <summary>
        /// Determines if a table contains any immediate child table within its cells.
        /// Does not recursively traverse through those tables to check for further tables.
        /// <returns>Returns true if at least one child cell contains a table.
        /// Returns false if no cells in the table contains a table.</returns>
        /// </summary>
        private static int GetChildTableCount(Table table)
        {
            int tableCount = 0;
            // Iterate through all child rows in the table
            foreach (Row row in table.Rows.OfType<Row>())
            {
                // Iterate through all child cells in the row
                foreach (Cell Cell in row.Cells.OfType<Cell>())
                {
                    // Retrieve the collection of child tables of this cell
                    TableCollection childTables = Cell.Tables;

                    // If this cell has a table as a child then return true
                    if (childTables.Count > 0)
                        tableCount++;
                }
            }

            // No cell contains a table
            return tableCount;
        }
        //ExEnd

        [Test]
        public void ConvertTextBoxToTable()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text box
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 50);

            // Move the builder into the text box and write text
            builder.MoveTo(textBox.LastParagraph);
            builder.Write("Hello world!");

            // Convert all shape nodes which contain child nodes
            // We convert the collection to an array as static "snapshot" because the original textboxes will be removed after conversion which will
            // invalidate the enumerator
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).ToArray().OfType<Shape>())
            {
                if (shape.HasChildNodes)
                {
                    ConvertTextboxToTable(shape);
                }
            }

            doc.Save(ArtifactsDir + "Table.ConvertTextBoxToTable.html");
        }

        /// <summary>
        /// Converts a textbox to a table by copying the same content and formatting.
        /// Currently export to HTML will render the textbox as an image which looses any text functionality.
        /// This is useful to convert textboxes in order to retain proper text.
        /// </summary>
        /// <param name="textBox">The textbox shape to convert to a table</param>
        private static void ConvertTextboxToTable(Shape textBox)
        {
            if (textBox.StoryType != StoryType.Textbox)
                throw new ArgumentException("Can only convert a shape of type textbox");

            Document doc = (Document) textBox.Document;
            Section section = (Section) textBox.GetAncestor(NodeType.Section);

            // Create a table to replace the textbox and transfer the same content and formatting
            Table table = new Table(doc);
            // Ensure that the table contains a row and a cell
            table.EnsureMinimum();
            // Use fixed column widths
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // A shape is inline level (within a paragraph) where a table can only be block level so insert the table
            // after the paragraph which contains the shape
            Node shapeParent = textBox.ParentNode;
            shapeParent.ParentNode.InsertAfter(table, shapeParent);

            // If the textbox is not inline then try to match the shape's left position using the table's left indent
            if (!textBox.IsInline && textBox.Left < section.PageSetup.PageWidth)
                table.LeftIndent = textBox.Left;

            // We are only using one cell to replicate a textbox so we can make use of the FirstRow and FirstCell property
            // Carry over borders and shading
            Row firstRow = table.FirstRow;
            Cell firstCell = firstRow.FirstCell;
            firstCell.CellFormat.Borders.Color = textBox.StrokeColor;
            firstCell.CellFormat.Borders.LineWidth = textBox.StrokeWeight;
            firstCell.CellFormat.Shading.BackgroundPatternColor = textBox.Fill.Color;

            // Transfer the same height and width of the textbox to the table
            firstRow.RowFormat.HeightRule = HeightRule.Exactly;
            firstRow.RowFormat.Height = textBox.Height;
            firstCell.CellFormat.Width = textBox.Width;
            table.AllowAutoFit = false;

            // Replicate the textbox's horizontal alignment
            TableAlignment horizontalAlignment;
            switch (textBox.HorizontalAlignment)
            {
                case HorizontalAlignment.Left:
                    horizontalAlignment = TableAlignment.Left;
                    break;
                case HorizontalAlignment.Center:
                    horizontalAlignment = TableAlignment.Center;
                    break;
                case HorizontalAlignment.Right:
                    horizontalAlignment = TableAlignment.Right;
                    break;
                default:
                    // Most other options are left by default
                    horizontalAlignment = TableAlignment.Left;
                    break;
            }

            table.Alignment = horizontalAlignment;
            firstCell.RemoveAllChildren();

            // Append all content from the textbox to the new table
            foreach (Node node in textBox.GetChildNodes(NodeType.Any, false).ToArray())
            {
                table.FirstRow.FirstCell.AppendChild(node);
            }

            // Remove the empty textbox from the document
            textBox.Remove();
        }

        [Test]
        public void EnsureTableMinimum()
        {
            //ExStart
            //ExFor:Table.EnsureMinimum
            //ExSummary:Shows how to ensure a table node is valid.
            Document doc = new Document();

            // Create a new table and add it to the document
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Ensure the table is valid (has at least one row with one cell)
            table.EnsureMinimum();
            //ExEnd
        }

        [Test]
        public void EnsureRowMinimum()
        {
            //ExStart
            //ExFor:Row.EnsureMinimum
            //ExSummary:Shows how to ensure a row node is valid.
            Document doc = new Document();

            // Create a new table and add it to the document
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Create a new row and add it to the table
            Row row = new Row(doc);
            table.AppendChild(row);

            // Ensure the row is valid (has at least one cell)
            row.EnsureMinimum();
            //ExEnd
        }

        [Test]
        public void EnsureCellMinimum()
        {
            //ExStart
            //ExFor:Cell.EnsureMinimum
            //ExSummary:Shows how to ensure a cell node is valid.
            Document doc = new Document(MyDir + "Tables.docx");

            // Gets the first cell in the document
            Cell cell = (Cell) doc.GetChild(NodeType.Cell, 0, true);

            // Ensure the cell is valid (the last child is a paragraph)
            cell.EnsureMinimum();
            //ExEnd
        }

        [Test]
        public void SetOutlineBorders()
        {
            //ExStart
            //ExFor:Table.Alignment
            //ExFor:TableAlignment
            //ExFor:Table.ClearBorders
            //ExFor:Table.ClearShading
            //ExFor:Table.SetBorder
            //ExFor:TextureIndex
            //ExFor:Table.SetShading
            //ExSummary:Shows how to apply a outline border to a table.
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Align the table to the center of the page
            table.Alignment = TableAlignment.Center;

            // Clear any existing borders and shading from the table
            table.ClearBorders();
            table.ClearShading();

            // Set a green border around the table but not inside
            table.SetBorder(BorderType.Left, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Green, true);

            // Fill the cells with a light green solid color
            table.SetShading(TextureIndex.TextureSolid, Color.LightGreen, Color.Empty);

            doc.Save(ArtifactsDir + "Table.SetOutlineBorders.docx");
            //ExEnd

            // Verify the borders were set correctly
            Assert.AreEqual(TableAlignment.Center, table.Alignment);
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Top.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Left.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Right.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Bottom.Color.ToArgb());
            Assert.AreNotEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Horizontal.Color.ToArgb());
            Assert.AreNotEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Vertical.Color.ToArgb());
            Assert.AreEqual(Color.LightGreen.ToArgb(),
                table.FirstRow.FirstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb());
        }

        [Test]
        public void SetTableBorders()
        {
            //ExStart
            //ExFor:Table.SetBorders
            //ExSummary:Shows how to build a table with all borders enabled (grid).
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Clear any existing borders from the table
            table.ClearBorders();

            // Set a green border around and inside the table
            table.SetBorders(LineStyle.Single, 1.5, Color.Green);

            doc.Save(ArtifactsDir + "Table.SetAllBorders.doc");
            //ExEnd

            // Verify the borders were set correctly
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Top.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Left.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Right.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Bottom.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Horizontal.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Vertical.Color.ToArgb());
        }

        [Test]
        public void RowFormat()
        {
            //ExStart
            //ExFor:RowFormat
            //ExFor:Row.RowFormat
            //ExSummary:Shows how to modify formatting of a table row.
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Retrieve the first row in the table
            Row firstRow = table.FirstRow;

            // Modify some row level properties
            firstRow.RowFormat.Borders.LineStyle = LineStyle.None;
            firstRow.RowFormat.HeightRule = HeightRule.Auto;
            firstRow.RowFormat.AllowBreakAcrossPages = true;
            //ExEnd

            doc.Save(ArtifactsDir + "Table.RowFormat.doc");

            doc = new Document(ArtifactsDir + "Table.RowFormat.doc");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Assert.AreEqual(LineStyle.None, table.FirstRow.RowFormat.Borders.LineStyle);
            Assert.AreEqual(HeightRule.Auto, table.FirstRow.RowFormat.HeightRule);
            Assert.True(table.FirstRow.RowFormat.AllowBreakAcrossPages);
        }

        [Test]
        public void CellFormat()
        {
            //ExStart
            //ExFor:CellFormat
            //ExFor:Cell.CellFormat
            //ExSummary:Shows how to modify formatting of a table cell.
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Retrieve the first cell in the table
            Cell firstCell = table.FirstRow.FirstCell;

            // Modify some row level properties
            firstCell.CellFormat.Width = 30; // in points
            firstCell.CellFormat.Orientation = TextOrientation.Downward;
            firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen;
            //ExEnd

            doc.Save(ArtifactsDir + "Table.CellFormat.doc");

            doc = new Document(ArtifactsDir + "Table.CellFormat.doc");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Assert.AreEqual(30, table.FirstRow.FirstCell.CellFormat.Width);
            Assert.AreEqual(TextOrientation.Downward, table.FirstRow.FirstCell.CellFormat.Orientation);
            Assert.AreEqual(Color.LightGreen.ToArgb(),
                table.FirstRow.FirstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb());
        }

        [Test]
        public void GetDistance()
        {
            //ExStart
            //ExFor:Table.DistanceBottom
            //ExFor:Table.DistanceLeft
            //ExFor:Table.DistanceRight
            //ExFor:Table.DistanceTop
            //ExSummary:Shows the minimum distance operations between table boundaries and text.
            Document doc = new Document(MyDir + "Table wrapped by text.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(25.9d, table.DistanceTop);
            Assert.AreEqual(25.9d, table.DistanceBottom);
            Assert.AreEqual(17.3d, table.DistanceLeft);
            Assert.AreEqual(17.3d, table.DistanceRight);
            //ExEnd
        }

        [Test]
        public void ClearBorders()
        {
            //ExStart
            //ExFor:Table.ClearBorders
            //ExSummary:Shows how to remove all borders from a table.
            Document doc = new Document(MyDir + "Tables.docx");

            // Remove all borders from the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Clear the borders all cells in the table
            table.ClearBorders();

            doc.Save(ArtifactsDir + "Table.ClearBorders.doc");
            //ExEnd
        }

        [Test]
        public void ReplaceCellText()
        {
            //ExStart
            //ExFor:Range.Replace(String, String, FindReplaceOptions)
            //ExSummary:Shows how to replace all instances of String of text in a table and cell.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table and give it conditional styling on border colors based on the row being the first or last
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Carrots");
            builder.InsertCell();
            builder.Write("30");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Potatoes");
            builder.InsertCell();
            builder.Write("50");
            builder.EndTable();

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = true;
            options.FindWholeWordsOnly = true;

            // Replace any instances of our String in the entire table
            table.Range.Replace("Carrots", "Eggs", options);
            // Replace any instances of our String in the last cell of the table only
            table.LastRow.LastCell.Range.Replace("50", "20", options);

            doc.Save(ArtifactsDir + "Table.ReplaceCellText.doc");
            //ExEnd

            Assert.AreEqual("20", table.LastRow.LastCell.ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void PrintTableRange()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // The range text will include control characters such as "\a" for a cell
            // You can call ToString on the desired node to retrieve the plain text content

            // Print the plain text range of the table to the screen
            Console.WriteLine("Contents of the table: ");
            Console.WriteLine(table.Range.Text);
            
            // Print the contents of the second row to the screen
            Console.WriteLine("\nContents of the row: ");
            Console.WriteLine(table.Rows[1].Range.Text);

            // Print the contents of the last cell in the table to the screen
            Console.WriteLine("\nContents of the cell: ");
            Console.WriteLine(table.LastRow.LastCell.Range.Text);
            
            Assert.AreEqual("\aColumn 1\aColumn 2\aColumn 3\aColumn 4\a\a", table.Rows[1].Range.Text);
            Assert.AreEqual("Cell 12 contents\a", table.LastRow.LastCell.Range.Text);
        }

        [Test]
        public void CloneTable()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Create a clone of the table
            Table tableClone = (Table) table.Clone(true);

            // Insert the cloned table into the document after the original
            table.ParentNode.InsertAfter(tableClone, table);

            // Insert an empty paragraph between the two tables or else they will be combined into one
            // upon save. This has to do with document validation
            table.ParentNode.InsertAfter(new Paragraph(doc), table);

            doc.Save(ArtifactsDir + "Table.CloneTable.doc");
            
            // Verify that the table was cloned and inserted properly
            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(table.Range.Text, tableClone.Range.Text);

            foreach (Cell cell in tableClone.GetChildNodes(NodeType.Cell, true).OfType<Cell>())
                cell.RemoveAllChildren();
            
            Assert.AreEqual(string.Empty, tableClone.ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void DisableBreakAcrossPages()
        {
            //ExStart
            //ExFor:RowFormat.AllowBreakAcrossPages
            //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
            // Disable breaking across pages for all rows in the table
            Document doc = new Document(MyDir + "Table spanning two pages.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            foreach (Row row in table.OfType<Row>())
                row.RowFormat.AllowBreakAcrossPages = false;

            doc.Save(ArtifactsDir + "Table.DisableBreakAcrossPages.docx");
            //ExEnd

            Assert.False(table.FirstRow.RowFormat.AllowBreakAcrossPages);
            Assert.False(table.LastRow.RowFormat.AllowBreakAcrossPages);
        }

        [Test]
        public void AllowAutoFitOnTable()
        {
            Document doc = new Document();

            Table table = new Table(doc);
            table.EnsureMinimum();

            //ExStart
            //ExFor:Table.AllowAutoFit
            //ExSummary:Shows how to set a table to shrink or grow each cell to accommodate its contents.
            table.AllowAutoFit = true;
            //ExEnd
        }

        [Test]
        public void KeepTableTogether()
        {
            Document doc = new Document(MyDir + "Table spanning two pages.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            //ExStart
            //ExFor:ParagraphFormat.KeepWithNext
            //ExFor:Row.IsLastRow
            //ExFor:Paragraph.IsEndOfCell
            //ExFor:Paragraph.IsInCell
            //ExFor:Cell.ParentRow
            //ExFor:Cell.Paragraphs
            //ExSummary:Shows how to set a table to stay together on the same page.
            // To keep a table from breaking across a page we need to enable KeepWithNext 
            // for every paragraph in the table except for the last paragraphs in the last 
            // row of the table
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true).OfType<Cell>())
            foreach (Paragraph para in cell.Paragraphs.OfType<Paragraph>())
            {
                // Every paragraph that's inside a cell will have this flag set
                Assert.True(para.IsInCell);

                if (!(cell.ParentRow.IsLastRow && para.IsEndOfCell))
                    para.ParagraphFormat.KeepWithNext = true;
            }
            //ExEnd

            doc.Save(ArtifactsDir + "Table.KeepTableTogether.doc");

            // Verify the correct paragraphs were set properly
            foreach (Paragraph para in table.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
                if (para.IsEndOfCell && ((Cell) para.ParentNode).ParentRow.IsLastRow)
                    Assert.False(para.ParagraphFormat.KeepWithNext);
                else
                    Assert.True(para.ParagraphFormat.KeepWithNext);
        }

        [Test]
        public void AddClonedRowToTable()
        {
            //ExStart
            //ExFor:Row
            //ExSummary:Shows how to make a clone of the last row of a table and append it to the table.
            Document doc = new Document(MyDir + "Tables.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            // Clone the last row in the table
            Row clonedRow = (Row) table.LastRow.Clone(true);

            // Remove all content from the cloned row's cells. This makes the row ready for
            // new content to be inserted into
            foreach (Cell cell in clonedRow.Cells.OfType<Cell>())
                cell.RemoveAllChildren();

            // Add the row to the end of the table
            table.AppendChild(clonedRow);

            doc.Save(ArtifactsDir + "Table.AddCloneRowToTable.doc");
            //ExEnd

            // Verify that the row was cloned and appended properly
            Assert.AreEqual(6, table.Rows.Count);
            Assert.AreEqual(string.Empty, table.LastRow.ToString(SaveFormat.Text).Trim());
            Assert.AreEqual(5, table.LastRow.Cells.Count);
        }

        [Test]
        public void FixDefaultTableWidthsInAw105()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Keep a reference to the table being built
            Table table = builder.StartTable();

            // Apply some formatting
            builder.CellFormat.Width = 100;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Red;

            builder.InsertCell();
            // This will cause the table to be structured using column widths as in previous versions
            // instead of fitted to the page width like in the newer versions
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Continue with building your table as usual...
        }

        [Test]
        public void FixDefaultTableBordersIn105()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Keep a reference to the table being built
            Table table = builder.StartTable();

            builder.InsertCell();
            // Clear all borders to match the defaults used in previous versions
            table.ClearBorders();

            // Continue with building your table as usual...
        }

        [Test]
        public void FixDefaultTableFormattingExceptionIn105()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Keep a reference to the table being built
            Table table = builder.StartTable();

            // We must first insert a new cell which in turn inserts a row into the table
            builder.InsertCell();
            // Once a row exists in our table we can apply table wide formatting
            table.AllowAutoFit = true;

            // Continue with building your table as usual...
        }

        [Test]
        public void FixRowFormattingNotAppliedIn105()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();

            // For the first row this will be set correctly
            builder.RowFormat.HeadingFormat = true;

            builder.InsertCell();
            builder.Writeln("Text");
            builder.InsertCell();
            builder.Writeln("Text");

            // End the first row
            builder.EndRow();

            // Here we would normally define some other row formatting, such as disabling the 
            // heading format. However at the moment this will be ignored and the value from the 
            // first row reapplied to the row

            builder.InsertCell();

            // Instead make sure to specify the row formatting for the second row here
            builder.RowFormat.HeadingFormat = false;

            // Continue with building your table as usual...
        }

        [Test]
        public void GetIndexOfTableElements()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            //ExStart
            //ExFor:NodeCollection.IndexOf(Node)
            //ExSummary:Retrieves the index of a table in the document.
            NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
            int tableIndex = allTables.IndexOf(table);

            Row row = table.Rows[2];
            int rowIndex = table.IndexOf(row);

            Cell cell = row.LastCell;
            int cellIndex = row.IndexOf(cell);
            //ExEnd

            Assert.AreEqual(0, tableIndex);
            Assert.AreEqual(2, rowIndex);
            Assert.AreEqual(4, cellIndex);
        }

        [Test]
        public void GetPreferredWidthTypeAndValue()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            // Find the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            //ExStart
            //ExFor:PreferredWidthType
            //ExFor:PreferredWidth.Type
            //ExFor:PreferredWidth.Value
            //ExSummary:Retrieves the preferred width type of a table cell.
            Cell firstCell = table.FirstRow.FirstCell;
            PreferredWidthType type = firstCell.CellFormat.PreferredWidth.Type;
            double value = firstCell.CellFormat.PreferredWidth.Value;
            //ExEnd

            Assert.AreEqual(PreferredWidthType.Percent, type);
            Assert.AreEqual(11.16, value);
        }

        [Test]
        public void InsertTableUsingNodes()
        {
            //ExStart
            //ExFor:Table.AllowCellSpacing
            //ExFor:Row
            //ExFor:Row.RowFormat
            //ExFor:RowFormat
            //ExFor:Cell.CellFormat
            //ExFor:CellFormat
            //ExFor:CellFormat.Shading
            //ExFor:Cell.FirstParagraph
            //ExSummary:Shows how to insert a table using the constructors of nodes.
            Document doc = new Document();

            // We start by creating the table object. Note how we must pass the document object
            // to the constructor of each node. This is because every node we create must belong
            // to some document
            Table table = new Table(doc);
            // Add the table to the document
            doc.FirstSection.Body.AppendChild(table);

            // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
            // to ensure that the specified node is valid, in this case a valid table should have at least one
            // row and one cell, therefore this method creates them for us

            // Instead we will handle creating the row and table ourselves. This would be the best way to do this
            // if we were creating a table inside an algorithm for example
            Row row = new Row(doc);
            row.RowFormat.AllowBreakAcrossPages = true;
            table.AppendChild(row);

            // We can now apply any auto fit settings
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Create a cell and add it to the row
            Cell cell = new Cell(doc);
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            cell.CellFormat.Width = 80;

            // Add a paragraph to the cell as well as a new run with some text
            cell.AppendChild(new Paragraph(doc));
            cell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 1 Text"));

            // Add the cell to the row
            row.AppendChild(cell);

            // We would then repeat the process for the other cells and rows in the table
            // We can also speed things up by cloning existing cells and rows
            row.AppendChild(cell.Clone(false));
            row.LastCell.AppendChild(new Paragraph(doc));
            row.LastCell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 2 Text"));

            // Remove spacing between cells
            table.AllowCellSpacing = false;

            doc.Save(ArtifactsDir + "Table.InsertTableUsingNodes.doc");
            //ExEnd

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Row, true).Count);
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual("Row 1, Cell 1 Text\r\nRow 1, Cell 2 Text",
                doc.FirstSection.Body.Tables[0].ToString(SaveFormat.Text).Trim());
        }

        //ExStart
        //ExFor:Table
        //ExFor:Row
        //ExFor:Cell
        //ExFor:Table.#ctor(DocumentBase)
        //ExFor:Table.Title
        //ExFor:Table.Description
        //ExFor:Row.#ctor(DocumentBase)
        //ExFor:Cell.#ctor(DocumentBase)
        //ExSummary:Shows how to build a nested table without using DocumentBuilder.
        [Test] //ExSkip
        public void CreateNestedTable()
        {
            Document doc = new Document();

            // Create the outer table with three rows and four columns
            Table outerTable = CreateTable(doc, 3, 4, "Outer Table");
            // Add it to the document body
            doc.FirstSection.Body.AppendChild(outerTable);

            // Create another table with two rows and two columns
            Table innerTable = CreateTable(doc, 2, 2, "Inner Table");
            // Add this table to the first cell of the outer table
            outerTable.FirstRow.FirstCell.AppendChild(innerTable);

            doc.Save(ArtifactsDir + "Table.CreateNestedTable.doc");

            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count); // ExSkip
            Assert.AreEqual(1, outerTable.FirstRow.FirstCell.Tables.Count); //ExSkip
            Assert.AreEqual(16, outerTable.GetChildNodes(NodeType.Cell, true).Count); //ExSkip
            Assert.AreEqual(4, innerTable.GetChildNodes(NodeType.Cell, true).Count); //ExSkip
            Assert.AreEqual("Aspose table title", innerTable.Title); //ExSkip
            Assert.AreEqual("Aspose table description", innerTable.Description); //ExSkip
        }

        /// <summary>
        /// Creates a new table in the document with the given dimensions and text in each cell.
        /// </summary>
        private static Table CreateTable(Document doc, int rowCount, int cellCount, string cellText)
        {
            Table table = new Table(doc);

            // Create the specified number of rows
            for (int rowId = 1; rowId <= rowCount; rowId++)
            {
                Row row = new Row(doc);
                table.AppendChild(row);

                // Create the specified number of cells for each row
                for (int cellId = 1; cellId <= cellCount; cellId++)
                {
                    Cell cell = new Cell(doc);
                    row.AppendChild(cell);
                    // Add a blank paragraph to the cell
                    cell.AppendChild(new Paragraph(doc));

                    // Add the text
                    cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                }
            }

            // You can add title and description to your table only when added at least one row to the table first
            // This properties are meaningful for ISO / IEC 29500 compliant DOCX documents(see the OoxmlCompliance class)
            // When saved to pre-ISO/IEC 29500 formats, the properties are ignored
            table.Title = "Aspose table title";
            table.Description = "Aspose table description";

            return table;
        }
        //ExEnd

        //ExStart
        //ExFor:CellFormat.HorizontalMerge
        //ExFor:CellFormat.VerticalMerge
        //ExFor:CellMerge
        //ExSummary:Prints the horizontal and vertical merge type of a cell.
        [Test] //ExSkip
        public void CheckCellsMerged()
        {
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            // Retrieve the first table in the document
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            foreach (Row row in table.Rows.OfType<Row>())
            {
                foreach (Cell cell in row.Cells.OfType<Cell>())
                {
                    Console.WriteLine(PrintCellMergeType(cell));
                }
            }

            Assert.AreEqual("The cell at R1, C1 is vertically merged",
                PrintCellMergeType(table.FirstRow.FirstCell)); //ExSkip
        }

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

            return isVerticallyMerged ? $"The cell at {cellLocation} is vertically merged" : $"The cell at {cellLocation} is not merged";
        }
        //ExEnd

        [Test]
        public void MergeCellRange()
        {
            // Open the document
            Document doc = new Document(MyDir + "Tables.docx");

            // Retrieve the first table in the body of the first section
            Table table = doc.FirstSection.Body.Tables[0];

            // We want to merge the range of cells found in between these two cells
            Cell cellStartRange = table.Rows[2].Cells[2];
            Cell cellEndRange = table.Rows[3].Cells[3];

            // Merge all the cells between the two specified cells into one
            MergeCells(cellStartRange, cellEndRange);

            // Save the document
            doc.Save(ArtifactsDir + "Table.MergeCellRange.doc");

            // Verify the cells were merged
            int mergedCellsCount = 0;
            foreach (Node node in table.GetChildNodes(NodeType.Cell, true))
            {
                Cell cell = (Cell) node;
                if (cell.CellFormat.HorizontalMerge != CellMerge.None ||
                    cell.CellFormat.VerticalMerge != CellMerge.None)
                    mergedCellsCount++;
            }

            Assert.AreEqual(4, mergedCellsCount);
            Assert.True(table.Rows[2].Cells[2].CellFormat.HorizontalMerge == CellMerge.First);
            Assert.True(table.Rows[2].Cells[2].CellFormat.VerticalMerge == CellMerge.First);
            Assert.True(table.Rows[3].Cells[3].CellFormat.HorizontalMerge == CellMerge.Previous);
            Assert.True(table.Rows[3].Cells[3].CellFormat.VerticalMerge == CellMerge.Previous);
        }

        /// <summary>
        /// Merges the range of cells found between the two specified cells both horizontally and vertically. Can span over multiple rows.
        /// </summary>
        public static void MergeCells(Cell startCell, Cell endCell)
        {
            Table parentTable = startCell.ParentRow.ParentTable;

            // Find the row and cell indices for the start and end cell
            Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell),
                parentTable.IndexOf(startCell.ParentRow));
            Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow));
            // Create the range of cells to be merged based off these indices
            // Inverse each index if the end cell if before the start cell
            Rectangle mergeRange = new Rectangle(
                Math.Min(startCellPos.X, endCellPos.X),
                Math.Min(startCellPos.Y, endCellPos.Y),
                Math.Abs(endCellPos.X - startCellPos.X) + 1,
                Math.Abs(endCellPos.Y - startCellPos.Y) + 1);

            foreach (Row row in parentTable.Rows.OfType<Row>())
            {
                foreach (Cell cell in row.Cells.OfType<Cell>())
                {
                    Point currentPos = new Point(row.IndexOf(cell), parentTable.IndexOf(row));
                    // Check if the current cell is inside our merge range then merge it
                    if (mergeRange.Contains(currentPos))
                    {
                        cell.CellFormat.HorizontalMerge =
                            currentPos.X == mergeRange.X ? CellMerge.First : CellMerge.Previous;
                        cell.CellFormat.VerticalMerge =
                            currentPos.Y == mergeRange.Y ? CellMerge.First : CellMerge.Previous;
                    }
                }
            }
        }

        [Test]
        public void CombineTables()
        {
            //ExStart
            //ExFor:Cell.CellFormat
            //ExFor:CellFormat.Borders
            //ExFor:Table.Rows
            //ExFor:Table.FirstRow
            //ExFor:CellFormat.ClearFormatting
            //ExSummary:Shows how to combine the rows from two tables into one.
            // Load the document
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first and second table in the document
            // The rows from the second table will be appended to the end of the first table
            Table firstTable = (Table) doc.GetChild(NodeType.Table, 0, true);
            Table secondTable = (Table) doc.GetChild(NodeType.Table, 1, true);

            // Append all rows from the current table to the next
            // Due to the design of tables even tables with different cell count and widths can be joined into one table
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the empty table container
            secondTable.Remove();

            doc.Save(ArtifactsDir + "Table.CombineTables.doc");
            //ExEnd

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(9, doc.FirstSection.Body.Tables[0].Rows.Count);
            Assert.AreEqual(42, doc.FirstSection.Body.Tables[0].GetChildNodes(NodeType.Cell, true).Count);
        }

        [Test]
        public void SplitTable()
        {
            // Load the document
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first table in the document
            Table firstTable = (Table) doc.GetChild(NodeType.Table, 0, true);

            // We will split the table at the third row (inclusive)
            Row row = firstTable.Rows[2];

            // Create a new container for the split table
            Table table = (Table) firstTable.Clone(false);

            // Insert the container after the original
            firstTable.ParentNode.InsertAfter(table, firstTable);

            // Add a buffer paragraph to ensure the tables stay apart
            firstTable.ParentNode.InsertAfter(new Paragraph(doc), firstTable);

            Row currentRow;

            do
            {
                currentRow = firstTable.LastRow;
                table.PrependChild(currentRow);
            } while (currentRow != row);

            doc.Save(ArtifactsDir + "Table.SplitTable.doc");

            doc = new Document(ArtifactsDir + "Table.SplitTable.doc");
            // Test we are adding the rows in the correct order and the 
            // selected row was also moved
            Assert.AreEqual(row, table.FirstRow);

            Assert.AreEqual(2, firstTable.Rows.Count);
            Assert.AreEqual(3, table.Rows.Count);
            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Table, true).Count);
        }

        [Test]
        public void CheckDefaultValuesForFloatingTableProperties()
        {
            //ExStart
            //ExFor:Table.TextWrapping
            //ExFor:TextWrapping
            //ExSummary:Shows how to work with table text wrapping.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = DocumentHelper.InsertTable(builder);

            if (table.TextWrapping == TextWrapping.Around)
            {
                Assert.AreEqual(HorizontalAlignment.Default, table.RelativeHorizontalAlignment);
                Assert.AreEqual(VerticalAlignment.Default, table.RelativeVerticalAlignment);
                Assert.AreEqual(RelativeHorizontalPosition.Column, table.HorizontalAnchor);
                Assert.AreEqual(RelativeVerticalPosition.Margin, table.VerticalAnchor);
                Assert.AreEqual(0, table.AbsoluteHorizontalDistance);
                Assert.AreEqual(0, table.AbsoluteVerticalDistance);
                Assert.AreEqual(true, table.AllowOverlap);
            }
            //ExEnd
        }

        [Test]
        public void GetFloatingTableProperties()
        {
            //ExStart
            //ExFor:Table.HorizontalAnchor
            //ExFor:Table.VerticalAnchor
            //ExFor:Table.AllowOverlap
            //ExFor:ShapeBase.AllowOverlap
            //ExSummary:Shows how get properties for floating tables
            Document doc = new Document(MyDir + "Table wrapped by text.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            if (table.TextWrapping == TextWrapping.Around)
            {
                Assert.AreEqual(RelativeHorizontalPosition.Margin, table.HorizontalAnchor);
                Assert.AreEqual(RelativeVerticalPosition.Paragraph, table.VerticalAnchor);
                Assert.AreEqual(false, table.AllowOverlap);
            }
            //ExEnd
        }

        [Test]
        public void ChangeFloatingTableProperties()
        {
            //ExStart
            //ExFor:Table.RelativeHorizontalAlignment
            //ExFor:Table.RelativeVerticalAlignment
            //ExFor:Table.AbsoluteHorizontalDistance
            //ExFor:Table.AbsoluteVerticalDistance
            //ExSummary:Shows how get/set properties for floating tables.
            Document doc = new Document(MyDir + "Table wrapped by text.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            table.AbsoluteHorizontalDistance = 10;
            table.AbsoluteVerticalDistance = 15;

            // Check that absolute distance was set correct
            Assert.AreEqual(10, table.AbsoluteHorizontalDistance);
            Assert.AreEqual(15, table.AbsoluteVerticalDistance);

            // Setting RelativeHorizontalAlignment will reset AbsoluteHorizontalDistance to default value and vice versa,
            // the same is for vertical positioning
            table.RelativeVerticalAlignment = VerticalAlignment.Top;
            table.RelativeHorizontalAlignment = HorizontalAlignment.Center;
            
            // Check that AbsoluteHorizontalDistance and AbsoluteVerticalDistance are reset 
            Assert.AreEqual(0, table.AbsoluteHorizontalDistance);
            Assert.AreEqual(0, table.AbsoluteVerticalDistance);
            Assert.AreEqual(VerticalAlignment.Top, table.RelativeVerticalAlignment);
            Assert.AreEqual(HorizontalAlignment.Center, table.RelativeHorizontalAlignment);

            doc.Save(ArtifactsDir + "Table.ChangeFloatingTableProperties.docx");
            //ExEnd
        }

        [Test]
        public void TableStyleCreation()
        {
            //ExStart
            //ExFor:Table.Bidi
            //ExFor:Table.CellSpacing
            //ExFor:Table.Style
            //ExFor:Table.StyleName
            //ExFor:TableStyle
            //ExFor:TableStyle.AllowBreakAcrossPages
            //ExFor:TableStyle.Bidi
            //ExFor:TableStyle.CellSpacing
            //ExFor:TableStyle.BottomPadding
            //ExFor:TableStyle.LeftPadding
            //ExFor:TableStyle.RightPadding
            //ExFor:TableStyle.TopPadding
            //ExFor:TableStyle.Shading
            //ExFor:TableStyle.Borders
            //ExSummary:Shows how to create your own style settings for the table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
 
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Name");
            builder.InsertCell();
            builder.Write("مرحبًا");
            builder.EndRow();
            builder.InsertCell();
            builder.InsertCell();
            builder.EndTable();
 
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.AllowBreakAcrossPages = true;
            tableStyle.Bidi = true;
            tableStyle.CellSpacing = 5.0;
            tableStyle.BottomPadding = 20.0;
            tableStyle.LeftPadding = 5;
            tableStyle.RightPadding = 10;
            tableStyle.TopPadding = 20.0;
            tableStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            tableStyle.Borders.Color = Color.Black;
            tableStyle.Borders.LineStyle = LineStyle.DotDash;

            table.Style = tableStyle;

            // Some Table attributes are linked to style variables
            Assert.AreEqual(true, table.Bidi);
            Assert.AreEqual(5.0, table.CellSpacing);
            Assert.AreEqual("MyTableStyle1", table.StyleName);

            doc.Save(ArtifactsDir + "Table.TableStyleCreation.docx");
            //ExEnd
        }

        [Test]
        public void SetTableAligment()
        {
            //ExStart
            //ExFor:TableStyle.Alignment
            //ExFor:TableStyle.LeftIndent
            //ExSummary:Shows how to set table position.
            Document doc = new Document();
 
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            // By default AW uses Alignment instead of LeftIndent
            // To set table position use
            tableStyle.Alignment = TableAlignment.Center;
            // or
            tableStyle.LeftIndent = 55.0;
            //ExEnd
        }

        [Test]
        public void WorkWithTableConditionalStyles()
        {
            //ExStart
            //ExFor:ConditionalStyle
            //ExFor:ConditionalStyle.Shading
            //ExFor:ConditionalStyle.Borders
            //ExFor:ConditionalStyle.ParagraphFormat
            //ExFor:ConditionalStyle.BottomPadding
            //ExFor:ConditionalStyle.LeftPadding
            //ExFor:ConditionalStyle.RightPadding
            //ExFor:ConditionalStyle.TopPadding
            //ExFor:ConditionalStyle.Font
            //ExFor:ConditionalStyle.Type
            //ExFor:ConditionalStyleCollection.GetEnumerator
            //ExFor:ConditionalStyleCollection.FirstRow
            //ExFor:ConditionalStyleCollection.LastRow
            //ExFor:ConditionalStyleCollection.LastColumn
            //ExFor:ConditionalStyleCollection.Count
            //ExFor:ConditionalStyleCollection
            //ExFor:ConditionalStyleCollection.BottomLeftCell
            //ExFor:ConditionalStyleCollection.BottomRightCell
            //ExFor:ConditionalStyleCollection.EvenColumnBanding
            //ExFor:ConditionalStyleCollection.EvenRowBanding
            //ExFor:ConditionalStyleCollection.FirstColumn
            //ExFor:ConditionalStyleCollection.Item(ConditionalStyleType)
            //ExFor:ConditionalStyleCollection.Item(TableStyleOverrideType)
            //ExFor:ConditionalStyleCollection.Item(Int32)
            //ExFor:ConditionalStyleCollection.OddColumnBanding
            //ExFor:ConditionalStyleCollection.OddRowBanding
            //ExFor:ConditionalStyleCollection.TopLeftCell
            //ExFor:ConditionalStyleCollection.TopRightCell
            //ExFor:ConditionalStyleType
            //ExFor:TableStyle.ConditionalStyles
            //ExSummary:Shows how to work with certain area styles of a table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table, which we will partially style
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1, to be formatted");
            builder.InsertCell();
            builder.Write("Cell 2, to be formatted");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Cell 3, to be left unformatted");
            builder.InsertCell();
            builder.Write("Cell 4, to be left unformatted");
            builder.EndTable();

            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            // There is a different ways how to get conditional styles:
            // by conditional style type
            tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.AliceBlue;
            // by index
            tableStyle.ConditionalStyles[0].Borders.Color = Color.Black;
            tableStyle.ConditionalStyles[0].Borders.LineStyle = LineStyle.DotDash;
            Assert.AreEqual(ConditionalStyleType.FirstRow, tableStyle.ConditionalStyles[0].Type);
            // directly from ConditionalStyleCollection
            tableStyle.ConditionalStyles.FirstRow.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            // To see this in Word document select Total Row checkbox in Design Tab
            tableStyle.ConditionalStyles.LastRow.BottomPadding = 10;
            tableStyle.ConditionalStyles.LastRow.LeftPadding = 10;
            tableStyle.ConditionalStyles.LastRow.RightPadding = 10;
            tableStyle.ConditionalStyles.LastRow.TopPadding = 10;
            // To see this in Word document select Last Column checkbox in Design Tab
            tableStyle.ConditionalStyles.LastColumn.Font.Bold = true;

            // List all possible style conditions
            using (IEnumerator<ConditionalStyle> enumerator = tableStyle.ConditionalStyles.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    ConditionalStyle currentStyle = enumerator.Current;
                    if (currentStyle != null) Console.WriteLine(currentStyle.Type);
                }
            }

            // Apply conditional style to the table and save
            table.Style = tableStyle;
            
            doc.Save(ArtifactsDir + "Table.WorkWithTableConditionalStyles.docx");
            //ExEnd
        }

        [Test]
        public void ClearTableStyleFormatting()
        {
            //ExStart
            //ExFor:ConditionalStyle.ClearFormatting
            //ExFor:ConditionalStyleCollection.ClearFormatting
            //ExSummary:Shows how to reset all table styles.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table and give it conditional styling on border colors based on the row being the first or last
            builder.StartTable();
            builder.InsertCell();
            builder.Write("First row");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Last row");
            builder.EndTable();

            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.ConditionalStyles.FirstRow.Borders.Color = Color.Red;
            tableStyle.ConditionalStyles.LastRow.Borders.Color = Color.Blue;

            // You can reset styles from the specific table area
            tableStyle.ConditionalStyles[0].ClearFormatting();
            Assert.AreEqual(Color.Empty, tableStyle.ConditionalStyles.FirstRow.Borders.Color);

            // Or clear all table styles
            tableStyle.ConditionalStyles.ClearFormatting();
            Assert.AreEqual(Color.Empty, tableStyle.ConditionalStyles.LastRow.Borders.Color);
            //ExEnd
        }

        [Test]
        public void WorkWithOddEvenRowColumnStyles()
        {
            //ExStart
            //ExFor:TableStyle.ColumnStripe
            //ExFor:TableStyle.RowStripe
            //ExSummary:Shows how to work with odd/even row/column styles.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table and give it conditional styling on border colors based on row number parity
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Odd row");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Even row");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Odd row");
            builder.EndTable();

            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.Borders.Color = Color.Black;
            tableStyle.Borders.LineStyle = LineStyle.DotDash;
            // Define our stripe through one column and row
            tableStyle.ColumnStripe = 1;
            tableStyle.RowStripe = 1;
            // Let's start from the first row and second column
            tableStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.AliceBlue;
            tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.AliceBlue;
            
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            table.Style = tableStyle;

            doc.Save(ArtifactsDir + "Table.WorkWithOddEvenRowColumnStyles.docx");
            //ExEnd
        }

        [Test]
        public void ConvertToHorizontallyMergedCells()
        {
            //ExStart
            //ExFor:Table.ConvertToHorizontallyMergedCells
            //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.HorizontalMerge.
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            // MS Word does not write merge flags anymore, they define merged cells by its width
            // So AW by default define only 5 cells in a row and all of it didn't have horizontal merge flag
            Table table = doc.FirstSection.Body.Tables[0];
            Row row = table.Rows[0];
            Assert.AreEqual(5, row.Cells.Count);

            // To resolve this inconvenience, we have added new public method to convert cells which are horizontally merged
            // by its width to the cell horizontally merged by flags. Thus now we have 7 cells and some of them have
            // horizontal merge value
            table.ConvertToHorizontallyMergedCells();
            row = table.Rows[0];
            Assert.AreEqual(7, row.Cells.Count);

            Assert.AreEqual(CellMerge.None, row.Cells[0].CellFormat.HorizontalMerge);
            Assert.AreEqual(CellMerge.First, row.Cells[1].CellFormat.HorizontalMerge);
            Assert.AreEqual(CellMerge.Previous, row.Cells[2].CellFormat.HorizontalMerge);
            Assert.AreEqual(CellMerge.None, row.Cells[3].CellFormat.HorizontalMerge);
            Assert.AreEqual(CellMerge.First, row.Cells[4].CellFormat.HorizontalMerge);
            Assert.AreEqual(CellMerge.Previous, row.Cells[5].CellFormat.HorizontalMerge);
            Assert.AreEqual(CellMerge.None, row.Cells[6].CellFormat.HorizontalMerge);
            //ExEnd
        }
    }
}