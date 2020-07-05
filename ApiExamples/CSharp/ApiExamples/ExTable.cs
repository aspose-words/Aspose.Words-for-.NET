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
        public void CreateTable()
        {
            //ExStart
            //ExFor:Table
            //ExFor:Row
            //ExFor:Cell
            //ExFor:Table.#ctor(DocumentBase)
            //ExSummary:Shows how to create a simple table.
            Document doc = new Document();

            // Tables are placed in the body of a document
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Tables contain rows, which contain cells,
            // which contain contents such as paragraphs, runs and even other tables
            // Calling table.EnsureMinimum will also make sure that a table has at least one row, cell and paragraph
            Row firstRow = new Row(doc);
            table.AppendChild(firstRow);

            Cell firstCell = new Cell(doc);
            firstRow.AppendChild(firstCell);

            Paragraph paragraph = new Paragraph(doc);
            firstCell.AppendChild(paragraph);

            Run run = new Run(doc, "Hello world!");
            paragraph.AppendChild(run);

            doc.Save(ArtifactsDir + "Table.CreateTable.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.CreateTable.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual(1, table.FirstRow.Cells.Count);
            Assert.AreEqual("Hello world!\a\a", table.GetText().Trim());
        }

        [Test]
        public void RowCellFormat()
        {
            //ExStart
            //ExFor:Row.RowFormat
            //ExFor:RowFormat
            //ExFor:Cell.CellFormat
            //ExFor:CellFormat
            //ExFor:CellFormat.Shading
            //ExSummary:Shows how to modify the format of rows and cells.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("City");
            builder.InsertCell();
            builder.Write("Country");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("London");
            builder.InsertCell();
            builder.Write("U.K.");
            builder.EndTable();

            // The appearance of rows and individual cells can be edited using the respective formatting objects
            RowFormat rowFormat = table.FirstRow.RowFormat;
            rowFormat.Height = 25;
            rowFormat.Borders[BorderType.Bottom].Color = Color.Red;

            CellFormat cellFormat = table.LastRow.FirstCell.CellFormat;
            cellFormat.Width = 100;
            cellFormat.Shading.BackgroundPatternColor = Color.Orange;

            doc.Save(ArtifactsDir + "Table.RowCellFormat.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.RowCellFormat.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual("City\aCountry\a\aLondon\aU.K.\a\a", table.GetText().Trim());

            rowFormat = table.FirstRow.RowFormat;

            Assert.AreEqual(25.0d, rowFormat.Height);
            Assert.AreEqual(Color.Red.ToArgb(), rowFormat.Borders[BorderType.Bottom].Color.ToArgb());

            cellFormat = table.LastRow.FirstCell.CellFormat;

            Assert.AreEqual(110.8d, cellFormat.Width);
            Assert.AreEqual(Color.Orange.ToArgb(), cellFormat.Shading.BackgroundPatternColor.ToArgb());
        }

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
            // which can contain block level nodes. For example, you can retrieve tables from header or from a cell
            // containing another table (nested tables)
            TableCollection tables = doc.FirstSection.Body.Tables;

            // We can make a new array to clone all the tables in the collection
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
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            Assert.AreEqual(5, tables.Count); //ExSkip

            for (int i = 0; i < tables.Count; i++)
            {
                Table table = (Table)tables[i];

                // Find out if any cells in the table have tables themselves as children
                int count = GetChildTableCount(table);
                Console.WriteLine("Table #{0} has {1} tables directly within its cells", i, count);

                // We can also do the opposite; finding out if the table is nested inside another table and at what depth
                int tableDepth = GetNestedDepthOfTable(table);

                if (tableDepth > 0)
                    Console.WriteLine("Table #{0} is nested inside another table at depth of {1}", i,
                        tableDepth);
                else
                    Console.WriteLine("Table #{0} is a non nested table (is not a child of another table)", i);
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
                // Every time we find a table a level up, we increase the depth counter and then try to find an
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
        /// Currently export to HTML will render the textbox as an image which loses any text functionality.
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

            // Currently, the table does not contain any rows, cells or nodes that can have content added to them
            Assert.AreEqual(0, table.GetChildNodes(NodeType.Any, true).Count);

            // This method ensures that the table has one row, one cell and one paragraph; the minimal nodes required to begin editing
            table.EnsureMinimum();
            table.FirstRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));
            //ExEnd

            Assert.AreEqual(4, table.GetChildNodes(NodeType.Any, true).Count);
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

            // Currently, the row does not contain any cells or nodes that can have content added to them
            Assert.AreEqual(0, row.GetChildNodes(NodeType.Any, true).Count);

            // Ensure the row has at least one cell with one paragraph that we can edit
            row.EnsureMinimum();
            row.FirstCell.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));
            //ExEnd

            Assert.AreEqual(3, row.GetChildNodes(NodeType.Any, true).Count);
        }

        [Test]
        public void EnsureCellMinimum()
        {
            //ExStart
            //ExFor:Cell.EnsureMinimum
            //ExSummary:Shows how to ensure a cell node is valid.
            Document doc = new Document();

            // Create a new table and add it to the document
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Create a new row with a new cell and append it to the table
            Row row = new Row(doc);
            table.AppendChild(row);

            Cell cell = new Cell(doc);
            row.AppendChild(cell);

            // Currently, the cell does not contain any cells or nodes that can have content added to them
            Assert.AreEqual(0, cell.GetChildNodes(NodeType.Any, true).Count);

            // Ensure the cell has at least one paragraph that we can edit
            cell.EnsureMinimum();
            cell.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));
            //ExEnd

            Assert.AreEqual(2, cell.GetChildNodes(NodeType.Any, true).Count);
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
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

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

            doc = new Document(ArtifactsDir + "Table.SetOutlineBorders.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(TableAlignment.Center, table.Alignment);

            BorderCollection borders = table.FirstRow.RowFormat.Borders;

            Assert.AreEqual(Color.Green.ToArgb(), borders.Top.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), borders.Left.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), borders.Right.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), borders.Bottom.Color.ToArgb());
            Assert.AreNotEqual(Color.Green.ToArgb(), borders.Horizontal.Color.ToArgb());
            Assert.AreNotEqual(Color.Green.ToArgb(), borders.Vertical.Color.ToArgb());
            Assert.AreEqual(Color.LightGreen.ToArgb(), table.FirstRow.FirstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb());
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

            doc.Save(ArtifactsDir + "Table.SetAllBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.SetAllBorders.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            
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

            doc.Save(ArtifactsDir + "Table.RowFormat.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.RowFormat.docx");
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

            doc.Save(ArtifactsDir + "Table.CellFormat.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.CellFormat.docx");

            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Assert.AreEqual(30, table.FirstRow.FirstCell.CellFormat.Width);
            Assert.AreEqual(TextOrientation.Downward, table.FirstRow.FirstCell.CellFormat.Orientation);
            Assert.AreEqual(Color.LightGreen.ToArgb(), table.FirstRow.FirstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb());
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

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(25.9d, table.DistanceTop);
            Assert.AreEqual(25.9d, table.DistanceBottom);
            Assert.AreEqual(17.3d, table.DistanceLeft);
            Assert.AreEqual(17.3d, table.DistanceRight);
            //ExEnd
        }

        [Test]
        public void Borders()
        {
            //ExStart
            //ExFor:Table.ClearBorders
            //ExSummary:Shows how to remove all borders from a table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Hello world!");
            builder.EndTable();

            // Set a color/thickness for the top border of the first row and verify the values
            Border topBorder = table.FirstRow.RowFormat.Borders[BorderType.Top];
            table.SetBorder(BorderType.Top, LineStyle.Double, 1.5, Color.Red, true);

            Assert.AreEqual(1.5d, topBorder.LineWidth);
            Assert.AreEqual(Color.Red.ToArgb(), topBorder.Color.ToArgb());
            Assert.AreEqual(LineStyle.Double, topBorder.LineStyle);

            // Clear the borders all cells in the table
            table.ClearBorders();
            doc.Save(ArtifactsDir + "Table.ClearBorders.docx");

            // Upon re-opening the saved document, the new border attributes can be verified
            doc = new Document(ArtifactsDir + "Table.ClearBorders.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            topBorder = table.FirstRow.RowFormat.Borders[BorderType.Top];

            Assert.AreEqual(0.0d, topBorder.LineWidth);
            Assert.AreEqual(Color.Empty.ToArgb(), topBorder.Color.ToArgb());
            Assert.AreEqual(LineStyle.None, topBorder.LineStyle);
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

            doc.Save(ArtifactsDir + "Table.ReplaceCellText.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.ReplaceCellText.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual("Eggs\a30\a\aPotatoes\a20\a\a", table.GetText().Trim());
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

            doc = new Document(ArtifactsDir + "Table.DisableBreakAcrossPages.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.False(table.FirstRow.RowFormat.AllowBreakAcrossPages);
            Assert.False(table.LastRow.RowFormat.AllowBreakAcrossPages);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AllowAutoFitOnTable(bool allowAutoFit)
        {
            //ExStart
            //ExFor:Table.AllowAutoFit
            //ExSummary:Shows how to set a table to shrink or grow each cell to accommodate its contents.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Write(
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.Write(
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.EndRow();
            builder.EndTable();

            table.AllowAutoFit = allowAutoFit;

            doc.Save(ArtifactsDir + "Table.AllowAutoFitOnTable.html");
            //ExEnd

            if (allowAutoFit)
            {
                TestUtil.FileContainsString(
                    "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">",
                    ArtifactsDir + "Table.AllowAutoFitOnTable.html");
                TestUtil.FileContainsString(
                    "<td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">",
                    ArtifactsDir + "Table.AllowAutoFitOnTable.html");
            }
            else
            {
                TestUtil.FileContainsString(
                    "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">",
                    ArtifactsDir + "Table.AllowAutoFitOnTable.html");
                TestUtil.FileContainsString(
                    "<td style=\"width:7.2pt; border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">",
                    ArtifactsDir + "Table.AllowAutoFitOnTable.html");
            }
        }

        [Test]
        public void KeepTableTogether()
        {
            //ExStart
            //ExFor:ParagraphFormat.KeepWithNext
            //ExFor:Row.IsLastRow
            //ExFor:Paragraph.IsEndOfCell
            //ExFor:Paragraph.IsInCell
            //ExFor:Cell.ParentRow
            //ExFor:Cell.Paragraphs
            //ExSummary:Shows how to set a table to stay together on the same page.
            Document doc = new Document(MyDir + "Table spanning two pages.docx");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Enabling KeepWithNext for every paragraph in the table except for the last ones in the last row
            // will prevent the table from being split across pages 
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true).OfType<Cell>())
                foreach (Paragraph para in cell.Paragraphs.OfType<Paragraph>())
                {
                    Assert.True(para.IsInCell);

                    if (!(cell.ParentRow.IsLastRow && para.IsEndOfCell))
                        para.ParagraphFormat.KeepWithNext = true;
                }

            doc.Save(ArtifactsDir + "Table.KeepTableTogether.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.KeepTableTogether.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            foreach (Paragraph para in table.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
                if (para.IsEndOfCell && ((Cell)para.ParentNode).ParentRow.IsLastRow)
                    Assert.False(para.ParagraphFormat.KeepWithNext);
                else
                    Assert.True(para.ParagraphFormat.KeepWithNext);
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
            // Once a row exists in our table, we can apply table wide formatting
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

            // Here we could define some other row formatting, such as disabling the heading format.
            // However, this will be ignored and the value from the first row reapplied to the row
            builder.InsertCell();

            // Instead make sure to specify the row formatting for the second row here
            builder.RowFormat.HeadingFormat = false;

            // Continue with building your table as usual...
        }

        [Test]
        public void GetIndexOfTableElements()
        {
            //ExStart
            //ExFor:NodeCollection.IndexOf(Node)
            //ExSummary:Shows how to get the indexes of nodes in the collections that contain them.
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);

            Assert.AreEqual(0, allTables.IndexOf(table));

            Row row = table.Rows[2];

            Assert.AreEqual(2, table.IndexOf(row));

            Cell cell = row.LastCell;

            Assert.AreEqual(4, row.IndexOf(cell));
            //ExEnd
        }

        [Test]
        public void GetPreferredWidthTypeAndValue()
        {
            //ExStart
            //ExFor:PreferredWidthType
            //ExFor:PreferredWidth.Type
            //ExFor:PreferredWidth.Value
            //ExSummary:Shows how to verify the preferred width type of a table cell.
            Document doc = new Document(MyDir + "Tables.docx");

            // Find the first table in the document
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Cell firstCell = table.FirstRow.FirstCell;

            Assert.AreEqual(PreferredWidthType.Percent, firstCell.CellFormat.PreferredWidth.Type);
            Assert.AreEqual(11.16, firstCell.CellFormat.PreferredWidth.Value);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AllowCellSpacing(bool allowCellSpacing)
        {
            //ExStart
            //ExFor:Table.AllowCellSpacing
            //ExFor:Table.CellSpacing
            //ExSummary:Shows how to enable spacing between individual cells in a table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Animal");
            builder.InsertCell();
            builder.Write("Class");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Dog");
            builder.InsertCell();
            builder.Write("Mammal");
            builder.EndTable();

            // Set the size of padding space between cells, and the switch that enables/negates this setting
            table.CellSpacing = 3;
            table.AllowCellSpacing = allowCellSpacing;

            doc.Save(ArtifactsDir + "Table.AllowCellSpacing.html");
            //ExEnd

            TestUtil.FileContainsString(
                allowCellSpacing
                    ? "<td style=\"border-style:solid; border-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">"
                    : "<td style=\"border-right-style:solid; border-right-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; " +
                      "padding-right:5.03pt; padding-left:5.03pt; vertical-align:top\">",
                ArtifactsDir + "Table.AllowCellSpacing.html");
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
        //ExFor:Cell.FirstParagraph
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

            doc.Save(ArtifactsDir + "Table.CreateNestedTable.docx");
            TestCreateNestedTable(new Document(ArtifactsDir + "Table.CreateNestedTable.docx")); //ExSkip
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
            // This properties are meaningful for ISO / IEC 29500 compliant .docx documents(see the OoxmlCompliance class)
            // When saved to pre-ISO/IEC 29500 formats, the properties are ignored
            table.Title = "Aspose table title";
            table.Description = "Aspose table description";

            return table;
        }
        //ExEnd

        private void TestCreateNestedTable(Document doc)
        {
            Table outerTable = (Table)doc.GetChild(NodeType.Table, 0, true);
            Table innerTable = (Table)doc.GetChild(NodeType.Table, 1, true);

            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(1, outerTable.FirstRow.FirstCell.Tables.Count);
            Assert.AreEqual(16, outerTable.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual(4, innerTable.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual("Aspose table title", innerTable.Title);
            Assert.AreEqual("Aspose table description", innerTable.Description);
        }

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
                foreach (Cell cell in row.Cells.OfType<Cell>())
                    Console.WriteLine(PrintCellMergeType(cell));
            
            Assert.AreEqual("The cell at R1, C1 is vertically merged", PrintCellMergeType(table.FirstRow.FirstCell)); //ExSkip
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
            //ExFor:CompositeNode.HasChildNodes
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

            doc.Save(ArtifactsDir + "Table.CombineTables.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.CombineTables.docx");

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

            doc.Save(ArtifactsDir + "Table.SplitTable.docx");

            doc = new Document(ArtifactsDir + "Table.SplitTable.docx");
            // Test we are adding the rows in the correct order and the 
            // selected row was also moved
            Assert.AreEqual(row, table.FirstRow);

            Assert.AreEqual(2, firstTable.Rows.Count);
            Assert.AreEqual(3, table.Rows.Count);
            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Table, true).Count);
        }

        [Test]
        public void WrapText()
        {
            //ExStart
            //ExFor:Table.TextWrapping
            //ExFor:TextWrapping
            //ExSummary:Shows how to work with table text wrapping.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table and a paragraph of text after it
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);

            builder.Font.Size = 16;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Set the table to wrap text around it and push it down into the paragraph below be setting the position
            table.TextWrapping = TextWrapping.Around;
            table.AbsoluteHorizontalDistance = 100;
            table.AbsoluteVerticalDistance = 20;

            doc.Save(ArtifactsDir + "Table.WrapText.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.WrapText.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(TextWrapping.Around, table.TextWrapping);
            Assert.AreEqual(100.0d, table.AbsoluteHorizontalDistance);
            Assert.AreEqual(20.0d, table.AbsoluteVerticalDistance);
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
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(TextWrapping.Around, table.TextWrapping);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, table.HorizontalAnchor);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, table.VerticalAnchor);
            Assert.AreEqual(false, table.AllowOverlap);
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
            //ExSummary:Shows how set the location of floating tables.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 1, cell 1");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);

            // We can set the table's location to a place on the page, such as the bottom right corner
            table.RelativeVerticalAlignment = VerticalAlignment.Bottom;
            table.RelativeHorizontalAlignment = HorizontalAlignment.Right;

            table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 2, cell 1");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);

            // We can also set a horizontal and vertical offset from the location in the paragraph where the table was inserted 
            table.AbsoluteVerticalDistance = 50;
            table.AbsoluteHorizontalDistance = 100;

            doc.Save(ArtifactsDir + "Table.ChangeFloatingTableProperties.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.ChangeFloatingTableProperties.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(VerticalAlignment.Bottom, table.RelativeVerticalAlignment);
            Assert.AreEqual(HorizontalAlignment.Right, table.RelativeHorizontalAlignment);

            table = (Table)doc.GetChild(NodeType.Table, 1, true);

            Assert.AreEqual(50.0d, table.AbsoluteVerticalDistance);
            Assert.AreEqual(100.0d, table.AbsoluteHorizontalDistance);
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
            //ExSummary:Shows how to create custom style settings for the table.
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
            tableStyle.CellSpacing = 5;
            tableStyle.BottomPadding = 20;
            tableStyle.LeftPadding = 5;
            tableStyle.RightPadding = 10;
            tableStyle.TopPadding = 20;
            tableStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            tableStyle.Borders.Color = Color.Blue;
            tableStyle.Borders.LineStyle = LineStyle.DotDash;

            table.Style = tableStyle;

            // Some Table attributes are linked to style variables
            Assert.True(table.Bidi);
            Assert.AreEqual(5.0d, table.CellSpacing);
            Assert.AreEqual("MyTableStyle1", table.StyleName);

            doc.Save(ArtifactsDir + "Table.TableStyleCreation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.TableStyleCreation.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.True(table.Bidi);
            Assert.AreEqual(5.0d, table.CellSpacing);
            Assert.AreEqual("MyTableStyle1", table.StyleName);
            Assert.AreEqual(0.0d, table.BottomPadding);
            Assert.AreEqual(0.0d, table.LeftPadding);
            Assert.AreEqual(0.0d, table.RightPadding);
            Assert.AreEqual(0.0d, table.TopPadding);
            Assert.AreEqual(6, table.FirstRow.RowFormat.Borders.Count(b => b.Color.ToArgb() == Color.Blue.ToArgb()));

            tableStyle = (TableStyle)doc.Styles["MyTableStyle1"];

            Assert.True(tableStyle.AllowBreakAcrossPages);
            Assert.True(tableStyle.Bidi);
            Assert.AreEqual(5.0d, tableStyle.CellSpacing);
            Assert.AreEqual(20.0d, tableStyle.BottomPadding);
            Assert.AreEqual(5.0d, tableStyle.LeftPadding);
            Assert.AreEqual(10.0d, tableStyle.RightPadding);
            Assert.AreEqual(20.0d, tableStyle.TopPadding);
            Assert.AreEqual(Color.AntiqueWhite.ToArgb(), tableStyle.Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.Blue.ToArgb(), tableStyle.Borders.Color.ToArgb());
            Assert.AreEqual(LineStyle.DotDash, tableStyle.Borders.LineStyle);
        }

        [Test]
        public void SetTableAlignment()
        {
            //ExStart
            //ExFor:TableStyle.Alignment
            //ExFor:TableStyle.LeftIndent
            //ExSummary:Shows how to set table position.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // There are two ways of horizontally aligning a table using a custom table style
            // One way is to align it to a location on the page, such as the center
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.Alignment = TableAlignment.Center;
            tableStyle.Borders.Color = Color.Blue;
            tableStyle.Borders.LineStyle = LineStyle.Single;

            // Insert a table and apply the style we created to it
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Aligned to the center of the page");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);
            
            table.Style = tableStyle;

            // We can also set a specific left indent to the style, and apply it to the table
            tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle2");
            tableStyle.LeftIndent = 55;
            tableStyle.Borders.Color = Color.Green;
            tableStyle.Borders.LineStyle = LineStyle.Single;

            table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Aligned according to left indent");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);

            table.Style = tableStyle;

            doc.Save(ArtifactsDir + "Table.TableStyleCreation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.TableStyleCreation.docx");

            tableStyle = (TableStyle)doc.Styles["MyTableStyle1"];

            Assert.AreEqual(TableAlignment.Center, tableStyle.Alignment);
            Assert.AreEqual(tableStyle, ((Table)doc.GetChild(NodeType.Table, 0, true)).Style);

            tableStyle = (TableStyle)doc.Styles["MyTableStyle2"];

            Assert.AreEqual(55.0d, tableStyle.LeftIndent);
            Assert.AreEqual(tableStyle, ((Table)doc.GetChild(NodeType.Table, 1, true)).Style);
        }

        [Test]
        public void ConditionalStyles()
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

            // Create a table
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Create a custom table style
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");

            // Conditional styles are formatting changes that affect only some of the cells of the table based on a predicate,
            // such as the cells being in the last row
            // We can access these conditional styles by style type like this
            tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.AliceBlue;

            // The same conditional style can be accessed by index
            tableStyle.ConditionalStyles[0].Borders.Color = Color.Black;
            tableStyle.ConditionalStyles[0].Borders.LineStyle = LineStyle.DotDash;
            Assert.AreEqual(ConditionalStyleType.FirstRow, tableStyle.ConditionalStyles[0].Type);

            // It can also be found in the ConditionalStyles collection as an attribute
            tableStyle.ConditionalStyles.FirstRow.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Apply padding and text formatting to conditional styles 
            tableStyle.ConditionalStyles.LastRow.BottomPadding = 10;
            tableStyle.ConditionalStyles.LastRow.LeftPadding = 10;
            tableStyle.ConditionalStyles.LastRow.RightPadding = 10;
            tableStyle.ConditionalStyles.LastRow.TopPadding = 10;
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

            // Apply conditional style to the table
            table.Style = tableStyle;

            // Changes to the first row are enabled by the table's style options be default,
            // but need to be manually enabled for some other parts, such as the last column/row
            table.StyleOptions = table.StyleOptions | TableStyleOptions.LastRow | TableStyleOptions.LastColumn;

            doc.Save(ArtifactsDir + "Table.ConditionalStyles.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.ConditionalStyles.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(TableStyleOptions.Default | TableStyleOptions.LastRow | TableStyleOptions.LastColumn, table.StyleOptions);
            ConditionalStyleCollection conditionalStyles = ((TableStyle)doc.Styles["MyTableStyle1"]).ConditionalStyles;

            Assert.AreEqual(ConditionalStyleType.FirstRow, conditionalStyles[0].Type);
            Assert.AreEqual(Color.AliceBlue.ToArgb(), conditionalStyles[0].Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.Black.ToArgb(), conditionalStyles[0].Borders.Color.ToArgb());
            Assert.AreEqual(LineStyle.DotDash, conditionalStyles[0].Borders.LineStyle);
            Assert.AreEqual(ParagraphAlignment.Center, conditionalStyles[0].ParagraphFormat.Alignment);

            Assert.AreEqual(ConditionalStyleType.LastRow, conditionalStyles[2].Type);
            Assert.AreEqual(10.0d, conditionalStyles[2].BottomPadding);
            Assert.AreEqual(10.0d, conditionalStyles[2].LeftPadding);
            Assert.AreEqual(10.0d, conditionalStyles[2].RightPadding);
            Assert.AreEqual(10.0d, conditionalStyles[2].TopPadding);

            Assert.AreEqual(ConditionalStyleType.LastColumn, conditionalStyles[3].Type);
            Assert.True(conditionalStyles[3].Font.Bold);
        }

        [Test]
        public void ClearTableStyleFormatting()
        {
            //ExStart
            //ExFor:ConditionalStyle.ClearFormatting
            //ExFor:ConditionalStyleCollection.ClearFormatting
            //ExSummary:Shows how to reset conditional table styles.
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

            // Conditional styles can be cleared for specific parts of the table 
            tableStyle.ConditionalStyles[0].ClearFormatting();
            Assert.AreEqual(Color.Empty, tableStyle.ConditionalStyles.FirstRow.Borders.Color);

            // Also, they can be cleared for the entire table
            tableStyle.ConditionalStyles.ClearFormatting();
            Assert.AreEqual(Color.Empty, tableStyle.ConditionalStyles.LastRow.Borders.Color);
            //ExEnd
        }

        [Test]
        public void AlternatingRowStyles()
        {
            //ExStart
            //ExFor:TableStyle.ColumnStripe
            //ExFor:TableStyle.RowStripe
            //ExSummary:Shows how to create conditional table styles that alternate between rows.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The conditional style of a table can be configured to apply a different color to the row/column,
            // based on whether the row/column is even or odd, creating an alternating color pattern
            // We can also apply a number n to the row/column banding, meaning that the color alternates after every n rows/columns instead of one 
            // Create a table where the columns will be banded by single columns and rows will banded in threes
            Table table = builder.StartTable();
            for (int i = 0; i < 15; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    builder.InsertCell();
                    builder.Writeln($"{(j % 2 == 0 ? "Even" : "Odd")} column.");
                    builder.Write($"Row banding {(i % 3 == 0 ? "start" : "continuation")}.");
                }
                builder.EndRow();
            }
            builder.EndTable();

            // Set a line style for all the borders of the table
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.Borders.Color = Color.Black;
            tableStyle.Borders.LineStyle = LineStyle.Double;

            // Set the two colors which will alternate over every 3 rows
            tableStyle.RowStripe = 3;
            tableStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;
            tableStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightCyan;

            // Set a color to apply to every even column, which will override any custom row coloring
            tableStyle.ColumnStripe = 1;
            tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightSalmon;

            // Apply the style to the table
            table.Style = tableStyle;

            // Row bands are automatically enabled, but column banding needs to be enabled manually like this
            // Row coloring will only be overridden if the column banding has been explicitly set a color
            table.StyleOptions = table.StyleOptions | TableStyleOptions.ColumnBands;

            doc.Save(ArtifactsDir + "Table.AlternatingRowStyles.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.AlternatingRowStyles.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            tableStyle = (TableStyle)doc.Styles["MyTableStyle1"];

            Assert.AreEqual(tableStyle, table.Style);
            Assert.AreEqual(table.StyleOptions | TableStyleOptions.ColumnBands, table.StyleOptions);

            Assert.AreEqual(Color.Black.ToArgb(), tableStyle.Borders.Color.ToArgb());
            Assert.AreEqual(LineStyle.Double, tableStyle.Borders.LineStyle);
            Assert.AreEqual(3, tableStyle.RowStripe);
            Assert.AreEqual(Color.LightBlue.ToArgb(), tableStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.LightCyan.ToArgb(), tableStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(1, tableStyle.ColumnStripe);
            Assert.AreEqual(Color.LightSalmon.ToArgb(), tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor.ToArgb());
        }

        [Test]
        public void ConvertToHorizontallyMergedCells()
        {
            //ExStart
            //ExFor:Table.ConvertToHorizontallyMergedCells
            //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.HorizontalMerge.
            Document doc = new Document(MyDir + "Table with merged cells.docx");

            // Microsoft Word does not write merge flags anymore; merged cells are defined by width instead.
            // So Aspose.Words by default defines only 5 cells in a row, and none of them have the horizontal merge flag.
            Table table = doc.FirstSection.Body.Tables[0];
            Row row = table.Rows[0];
            Assert.AreEqual(5, row.Cells.Count);

            // There is a public method to convert cells which are horizontally merged
            // by its width to the cell horizontally merged by flags.
            // Thus, we have 7 cells and some of them have horizontal merge value
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