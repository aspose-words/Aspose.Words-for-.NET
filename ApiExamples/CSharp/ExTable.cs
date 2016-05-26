// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;

using Aspose.Words;
using Aspose.Words.Drawing;
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
            //ExFor:Table
            //ExFor:Row.Cells
            //ExFor:Table.Rows
            //ExFor:Cell
            //ExFor:Row
            //ExFor:RowCollection
            //ExFor:CellCollection
            //ExFor:NodeCollection.IndexOf(Node)
            //ExSummary:Shows how to iterate through all tables in the document and display the content from each cell.
            Document doc = new Document(MyDir + "Table.Document.doc");

            // Here we get all tables from the Document node. You can do this for any other composite node
            // which can contain block level nodes. For example you can retrieve tables from header or from a cell
            // containing another table (nested tables).
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

            // Iterate through all tables in the document
            foreach (Table table in tables)
            {
                // Get the index of the table node as contained in the parent node of the table
                int tableIndex = table.ParentNode.ChildNodes.IndexOf(table);
                Console.WriteLine("Start of Table {0}", tableIndex);

                // Iterate through all rows in the table
                foreach (Row row in table.Rows)
                {
                    int rowIndex = table.Rows.IndexOf(row);
                    Console.WriteLine("\tStart of Row {0}", rowIndex);

                    // Iterate through all cells in the row
                    foreach (Cell cell in row.Cells)
                    {
                        int cellIndex = row.Cells.IndexOf(cell);
                        // Get the plain text content of this cell.
                        string cellText = cell.ToString(SaveFormat.Text).Trim();
                        // Print the content of the cell.
                        Console.WriteLine("\t\tContents of Cell:{0} = \"{1}\"", cellIndex, cellText);
                    }
                    //Console.WriteLine();
                    Console.WriteLine("\tEnd of Row {0}", rowIndex);
                }
                Console.WriteLine("End of Table {0}", tableIndex);
                Console.WriteLine();
            }
            //ExEnd

            Assert.Greater(tables.Count, 0);
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void CalcuateDepthOfNestedTablesCaller()
        {
            this.CalcuateDepthOfNestedTables();
        }

        //ExStart
        //ExFor:Node.GetAncestor(NodeType)
        //ExFor:Table.NodeType
        //ExFor:Cell.Tables
        //ExFor:TableCollection
        //ExFor:NodeCollection.Count
        //ExSummary:Shows how to find out if a table contains another table or if the table itself is nested inside another table.
        public void CalcuateDepthOfNestedTables()
        {
            Document doc = new Document(MyDir + "Table.NestedTables.doc");
            int tableIndex = 0;

            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                // First lets find if any cells in the table have tables themselves as children.
                int count = GetChildTableCount(table);
                Console.WriteLine("Table #{0} has {1} tables directly within its cells", tableIndex, count);

                // Now let's try the other way around, lets try find if the table is nested inside another table and at what depth.
                int tableDepth = GetNestedDepthOfTable(table);

                if (tableDepth > 0)
                    Console.WriteLine("Table #{0} is nested inside another table at depth of {1}", tableIndex, tableDepth);
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

            NodeType type = table.NodeType;
            // The parent of the table will be a Cell, instead attempt to find a grandparent that is of type Table
            Node parent = table.GetAncestor(type);

            while (parent != null)
            {
                // Every time we find a table a level up we increase the depth counter and then try to find an
                // ancestor of type table from the parent.
                depth++;
                parent = parent.GetAncestor(type);
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
            foreach (Row row in table.Rows)
            {
                // Iterate through all child cells in the row
                foreach (Cell Cell in row.Cells)
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

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void ConvertTextboxToTableCaller()
        {
            this.ConvertTextboxToTable();
        }

        //ExStart
        //ExId:TextboxToTable
        //ExSummary:Shows how to convert a textbox to a table and retain almost identical formatting. This is useful for HTML export.
        public void ConvertTextboxToTable()
        {
            // Open the document
            Document doc = new Document(MyDir + "Shape.Textbox.doc");

            // Convert all shape nodes which contain child nodes.
            // We convert the collection to an array as static "snapshot" because the original textboxes will be removed after conversion which will
            // invalidate the enumerator.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).ToArray())
            {
                if (shape.HasChildNodes)
                {
                    ConvertTextboxToTable(shape);
                }
            }

            doc.Save(MyDir + @"\Artifacts\Table.ConvertTextboxToTable.html");
        }

        /// <summary>
        /// Converts a textbox to a table by copying the same content and formatting.
        /// Currently export to HTML will render the textbox as an image which looses any text functionality.
        /// This is useful to convert textboxes in order to retain proper text.
        /// </summary>
        /// <param name="textbox">The textbox shape to convert to a table</param>
        private static void ConvertTextboxToTable(Shape textBox)
        {
            if (textBox.StoryType != StoryType.Textbox)
                throw new ArgumentException("Can only convert a shape of type textbox");

            Document doc = (Document)textBox.Document;
            Section section = (Section)textBox.GetAncestor(NodeType.Section);

            // Create a table to replace the textbox and transfer the same content and formatting.
            Table table = new Table(doc);
            // Ensure that the table contains a row and a cell.
            table.EnsureMinimum();
            // Use fixed column widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // A shape is inline level (within a paragraph) where a table can only be block level so insert the table
            // after the paragraph which contains the shape.
            Node shapeParent = textBox.ParentNode;
            shapeParent.ParentNode.InsertAfter(table, shapeParent);

            // If the textbox is not inline then try to match the shape's left position using the table's left indent.
            if (!textBox.IsInline && textBox.Left < section.PageSetup.PageWidth)
                table.LeftIndent = textBox.Left;

            // We are only using one cell to replicate a textbox so we can make use of the FirstRow and FirstCell property.
            // Carry over borders and shading.
            Row firstRow = table.FirstRow;
            Cell firstCell = firstRow.FirstCell;
            firstCell.CellFormat.Borders.Color = textBox.StrokeColor;
            firstCell.CellFormat.Borders.LineWidth = textBox.StrokeWeight;
            firstCell.CellFormat.Shading.BackgroundPatternColor = textBox.Fill.Color;

            // Transfer the same height and width of the textbox to the table.
            firstRow.RowFormat.HeightRule = HeightRule.Exactly;
            firstRow.RowFormat.Height = textBox.Height;
            firstCell.CellFormat.Width = textBox.Width;
            table.AllowAutoFit = false;

            // Replicate the textbox's horizontal alignment.
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
                    // Most other options are left by default.
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

            // Remove the empty textbox from the document.
            textBox.Remove();
        }
        //ExEnd

        [Test]
        public void EnsureTableMinimum()
        {
            //ExStart
            //ExFor:Table.EnsureMinimum
            //ExSummary:Shows how to ensure a table node is valid.
            Document doc = new Document();

            // Create a new table and add it to the document.
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Ensure the table is valid (has at least one row with one cell).
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

            // Create a new table and add it to the document.
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Create a new row and add it to the table.
            Row row = new Row(doc);
            table.AppendChild(row);

            // Ensure the row is valid (has at least one cell).
            row.EnsureMinimum();
            //ExEnd
        }

        [Test]
        public void EnsureCellMinimum()
        {
            //ExStart
            //ExFor:Cell.EnsureMinimum
            //ExSummary:Shows how to ensure a cell node is valid.
            Document doc = new Document(MyDir + "Table.Document.doc");

            // Gets the first cell in the document.
            Cell cell = (Cell)doc.GetChild(NodeType.Cell, 0, true);

            // Ensure the cell is valid (the last child is a paragraph).
            cell.EnsureMinimum();
            //ExEnd
        }

        [Test]
        public void SetTableBordersOutline()
        {
            //ExStart
            //ExFor:Table.Alignment
            //ExFor:TableAlignment
            //ExFor:Table.ClearBorders
            //ExFor:Table.SetBorder
            //ExFor:TextureIndex
            //ExFor:Table.SetShading
            //ExId:TableBordersOutline
            //ExSummary:Shows how to apply a outline border to a table.
            Document doc = new Document(MyDir + "Table.EmptyTable.doc");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Align the table to the center of the page.
            table.Alignment = TableAlignment.Center;

            // Clear any existing borders from the table.
            table.ClearBorders();

            // Set a green border around the table but not inside. 
            table.SetBorder(BorderType.Left, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Green, true);

            // Fill the cells with a light green solid color.
            table.SetShading(TextureIndex.TextureSolid, Color.LightGreen, Color.Empty);

            doc.Save(MyDir + @"\Artifacts\Table.SetOutlineBorders.doc");
            //ExEnd

            // Verify the borders were set correctly.
            doc = new Document(MyDir + @"\Artifacts\Table.SetOutlineBorders.doc");
            Assert.AreEqual(TableAlignment.Center, table.Alignment);
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Top.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Left.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Right.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Bottom.Color.ToArgb());
            Assert.AreNotEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Horizontal.Color.ToArgb());
            Assert.AreNotEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Vertical.Color.ToArgb());
            Assert.AreEqual(Color.LightGreen.ToArgb(), table.FirstRow.FirstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb());
        }

        [Test]
        public void SetTableBordersAll()
        {
            //ExStart
            //ExFor:Table.SetBorders
            //ExId:TableBordersAll
            //ExSummary:Shows how to build a table with all borders enabled (grid).
            Document doc = new Document(MyDir + "Table.EmptyTable.doc");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Clear any existing borders from the table.
            table.ClearBorders();

            // Set a green border around and inside the table.
            table.SetBorders(LineStyle.Single, 1.5, Color.Green);

            doc.Save(MyDir + @"\Artifacts\Table.SetAllBorders.doc");
            //ExEnd

            // Verify the borders were set correctly.
            doc = new Document(MyDir + @"\Artifacts\Table.SetAllBorders.doc");
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Top.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Left.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Right.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Bottom.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Horizontal.Color.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Vertical.Color.ToArgb());
        }

        [Test]
        public void RowFormatProperties()
        {
            //ExStart
            //ExFor:RowFormat
            //ExFor:Row.RowFormat
            //ExId:RowFormatProperties
            //ExSummary:Shows how to modify formatting of a table row.
            Document doc = new Document(MyDir + "Table.Document.doc");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Retrieve the first row in the table.
            Row firstRow = table.FirstRow;

            // Modify some row level properties.
            firstRow.RowFormat.Borders.LineStyle = LineStyle.None;
            firstRow.RowFormat.HeightRule = HeightRule.Auto;
            firstRow.RowFormat.AllowBreakAcrossPages = true;
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Table.RowFormat.doc");

            doc = new Document(MyDir + @"\Artifacts\Table.RowFormat.doc");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Assert.AreEqual(LineStyle.None, table.FirstRow.RowFormat.Borders.LineStyle);
            Assert.AreEqual(HeightRule.Auto, table.FirstRow.RowFormat.HeightRule);
            Assert.True(table.FirstRow.RowFormat.AllowBreakAcrossPages);
        }

        [Test]
        public void CellFormatProperties()
        {
            //ExStart
            //ExFor:CellFormat
            //ExFor:Cell.CellFormat
            //ExId:CellFormatProperties
            //ExSummary:Shows how to modify formatting of a table cell.
            Document doc = new Document(MyDir + "Table.Document.doc");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Retrieve the first cell in the table.
            Cell firstCell = table.FirstRow.FirstCell;

            // Modify some row level properties.
            firstCell.CellFormat.Width = 30; // in points
            firstCell.CellFormat.Orientation = TextOrientation.Downward;
            firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen;
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Table.CellFormat.doc");

            doc = new Document(MyDir + @"\Artifacts\Table.CellFormat.doc");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Assert.AreEqual(30, table.FirstRow.FirstCell.CellFormat.Width);
            Assert.AreEqual(TextOrientation.Downward, table.FirstRow.FirstCell.CellFormat.Orientation);
            Assert.AreEqual(Color.LightGreen.ToArgb(), table.FirstRow.FirstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb());
        }

        [Test]
        public void RemoveBordersFromAllCells()
        {
            //ExStart
            //ExFor:Table
            //ExFor:Table.ClearBorders
            //ExSummary:Shows how to remove all borders from a table.
            Document doc = new Document(MyDir + "Table.Document.doc");

            // Remove all borders from the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Clear the borders all cells in the table.
            table.ClearBorders();

            doc.Save(MyDir + @"\Artifacts\Table.ClearBorders.doc");
            //ExEnd
        }

        [Test]
        public void ReplaceTextInTable()
        {
            //ExStart
            //ExFor:Range.Replace(String, String, Boolean, Boolean)
            //ExFor:Cell
            //ExId:ReplaceTextTable
            //ExSummary:Shows how to replace all instances of string of text in a table and cell.
            Document doc = new Document(MyDir + "Table.SimpleTable.doc");

            // Get the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Replace any instances of our string in the entire table.
            table.Range.Replace("Carrots", "Eggs", true, true);
            // Replace any instances of our string in the last cell of the table only.
            table.LastRow.LastCell.Range.Replace("50", "20", true, true);

            doc.Save(MyDir + @"\Artifacts\Table.ReplaceCellText.doc");
            //ExEnd

            Assert.AreEqual("20", table.LastRow.LastCell.ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void PrintTableRange()
        {
            //ExStart
            //ExId:PrintTableRange
            //ExSummary:Shows how to print the text range of a table.
            Document doc = new Document(MyDir + "Table.SimpleTable.doc");

            // Get the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // The range text will include control characters such as "\a" for a cell.
            // You can call ToString on the desired node to retrieve the plain text content.

            // Print the plain text range of the table to the screen.
            Console.WriteLine("Contents of the table: ");
            Console.WriteLine(table.Range.Text);
            //ExEnd

            //ExStart
            //ExId:PrintRowAndCellRange
            //ExSummary:Shows how to print the text range of row and table elements.
            // Print the contents of the second row to the screen.
            Console.WriteLine("\nContents of the row: ");
            Console.WriteLine(table.Rows[1].Range.Text);

            // Print the contents of the last cell in the table to the screen.
            Console.WriteLine("\nContents of the cell: ");
            Console.WriteLine(table.LastRow.LastCell.Range.Text);
            //ExEnd

            Assert.AreEqual("Apples\r" + ControlChar.Cell + "20\r" + ControlChar.Cell + ControlChar.Cell, table.Rows[1].Range.Text);
            Assert.AreEqual("50\r\a", table.LastRow.LastCell.Range.Text);
        }

        [Test]
        public void CloneTable()
        {
            //ExStart
            //ExId:CloneTable
            //ExSummary:Shows how to make a clone of a table in the document and insert it after the original table.
            Document doc = new Document(MyDir + "Table.SimpleTable.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Create a clone of the table.
            Table tableClone = (Table)table.Clone(true);
            
            // Insert the cloned table into the document after the original
            table.ParentNode.InsertAfter(tableClone, table);

            // Insert an empty paragraph between the two tables or else they will be combined into one
            // upon save. This has to do with document validation.
            table.ParentNode.InsertAfter(new Paragraph(doc), table);

            doc.Save(MyDir + @"\Artifacts\Table.CloneTableAndInsert.doc");
            //ExEnd

            // Verify that the table was cloned and inserted properly.
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(table.Range.Text, tableClone.Range.Text);

            //ExStart
            //ExId:CloneTableRemoveContent
            //ExSummary:Shows how to remove all content from the cells of a cloned table.
            foreach (Cell cell in tableClone.GetChildNodes(NodeType.Cell, true))
                cell.RemoveAllChildren();
            //ExEnd

            Assert.AreEqual(String.Empty, tableClone.ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void RowFormatDisableBreakAcrossPages()
        {
            Document doc = new Document(MyDir + "Table.TableAcrossPage.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            //ExStart
            //ExFor:RowFormat.AllowBreakAcrossPages
            //ExId:RowFormatAllowBreaks
            //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
            // Disable breaking across pages for all rows in the table.
            foreach(Row row in table)
                row.RowFormat.AllowBreakAcrossPages = false;
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Table.DisableBreakAcrossPages.doc");

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
            //ExId:AllowAutoFit
            //ExSummary:Shows how to set a table to shrink or grow each cell to accommodate its contents.
            table.AllowAutoFit = true;
            //ExEnd
        }

        [Test]
        public void KeepTableTogether()
        {
            Document doc = new Document(MyDir + "Table.TableAcrossPage.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            //ExStart
            //ExFor:ParagraphFormat.KeepWithNext
            //ExFor:Row.IsLastRow
            //ExFor:Paragraph.IsEndOfCell
            //ExFor:Cell.ParentRow
            //ExFor:Cell.Paragraphs
            //ExId:KeepTableTogether
            //ExSummary:Shows how to set a table to stay together on the same page.
            // To keep a table from breaking across a page we need to enable KeepWithNext 
            // for every paragraph in the table except for the last paragraphs in the last 
            // row of the table.
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
                foreach (Paragraph para in cell.Paragraphs)
                    if (!(cell.ParentRow.IsLastRow && para.IsEndOfCell))
                        para.ParagraphFormat.KeepWithNext = true;
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Table.KeepTableTogether.doc");

            // Verify the correct paragraphs were set properly.
            foreach (Paragraph para in table.GetChildNodes(NodeType.Paragraph, true))
                if (para.IsEndOfCell && ((Cell)para.ParentNode).ParentRow.IsLastRow)
                    Assert.False(para.ParagraphFormat.KeepWithNext);
                else
                    Assert.True(para.ParagraphFormat.KeepWithNext);
        }

        [Test]
        public void AddClonedRowToTable()
        {
            //ExStart
            //ExFor:Row
            //ExId:AddClonedRowToTable
            //ExSummary:Shows how to make a clone of the last row of a table and append it to the table.
            Document doc = new Document(MyDir + "Table.SimpleTable.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Clone the last row in the table.
            Row clonedRow = (Row)table.LastRow.Clone(true);

            // Remove all content from the cloned row's cells. This makes the row ready for
            // new content to be inserted into.
            foreach (Cell cell in clonedRow.Cells)
                cell.RemoveAllChildren();

            // Add the row to the end of the table.
            table.AppendChild(clonedRow);

            doc.Save(MyDir + @"\Artifacts\Table.AddCloneRowToTable.doc");
            //ExEnd

            // Verify that the row was cloned and appended properly.
            Assert.AreEqual(5, table.Rows.Count);
            Assert.AreEqual(string.Empty, table.LastRow.ToString(SaveFormat.Text).Trim());
            Assert.AreEqual(2, table.LastRow.Cells.Count);
        }

        [Test]
        public void FixDefaultTableWidthsInAw105()
        {
            //ExStart
            //ExId:FixTablesDefaultFixedColumnWidth
            //ExSummary:Shows how to revert the default behaviour of table sizing to use column widths.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Keep a reference to the table being built.
            Table table = builder.StartTable();

            // Apply some formatting.
            builder.CellFormat.Width = 100;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Red;

            builder.InsertCell();
            // This will cause the table to be structured using column widths as in previous verisons
            // instead of fitted to the page width like in the newer versions.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths); 

            // Continue with building your table as usual...
            //ExEnd
        }

        [Test]
        public void FixDefaultTableBordersIn105()
        {
            //ExStart
            //ExId:FixTablesDefaultBorders
            //ExSummary:Shows how to revert the default borders on tables back to no border lines.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Keep a reference to the table being built.
            Table table = builder.StartTable();
  
            builder.InsertCell();
            // Clear all borders to match the defaults used in previous versions.
            table.ClearBorders();

            // Continue with building your table as usual...
            //ExEnd
        }

        [Test]
        public void FixDefaultTableFormattingExceptionIn105()
        {
            //ExStart
            //ExId:FixTableFormattingException
            //ExSummary:Shows how to avoid encountering an exception when applying table formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Keep a reference to the table being built.
            Table table = builder.StartTable();

            // We must first insert a new cell which in turn inserts a row into the table.
            builder.InsertCell();
            // Once a row exists in our table we can apply table wide formatting.
            table.AllowAutoFit = true;

            // Continue with building your table as usual...
            //ExEnd
        }

        [Test]
        public void FixRowFormattingNotAppliedIn105()
        {
            //ExStart
            //ExId:FixRowFormattingNotApplied
            //ExSummary:Shows how to fix row formatting not being applied to some rows.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();

            // For the first row this will be set correctly.
            builder.RowFormat.HeadingFormat = true;

            builder.InsertCell();
            builder.Writeln("Text");
            builder.InsertCell();
            builder.Writeln("Text");

            // End the first row.
            builder.EndRow();

            // Here we would normally define some other row formatting, such as disabling the 
            // heading format. However at the moment this will be ignored and the value from the 
            // first row reapplied to the row.

            builder.InsertCell();

            // Instead make sure to specify the row formatting for the second row here.
            builder.RowFormat.HeadingFormat = false;

            // Continue with building your table as usual...
            //ExEnd
        }

        [Test]
        public void GetIndexOfTableElements()
        {
            Document doc = new Document(MyDir + "Table.Document.doc");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            //ExStart
            //ExFor:NodeCollection.IndexOf
            //ExId:IndexOfTable
            //ExSummary:Retrieves the index of a table in the document.
            NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
            int tableIndex = allTables.IndexOf(table);
            //ExEnd

            Row row = table.Rows[2];
            //ExStart
            //ExFor:Row
            //ExFor:CompositeNode.IndexOf
            //ExId:IndexOfRow
            //ExSummary:Retrieves the index of a row in a table.
            int rowIndex = table.IndexOf(row);
            //ExEnd

            Cell cell = row.LastCell;
            //ExStart
            //ExFor:Cell
            //ExFor:CompositeNode.IndexOf
            //ExId:IndexOfCell
            //ExSummary:Retrieves the index of a cell in a row.
            int cellIndex = row.IndexOf(cell);
            //ExEnd

            Assert.AreEqual(0, tableIndex);
            Assert.AreEqual(2, rowIndex);
            Assert.AreEqual(4, cellIndex);
        }

        [Test]
        public void GetPreferredWidthTypeAndValue()
        {
            Document doc = new Document(MyDir + "Table.Document.doc");

            // Find the first table in the document
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            //ExStart
            //ExFor:PreferredWidthType
            //ExFor:PreferredWidth.Type
            //ExFor:PreferredWidth.Value
            //ExId:GetPreferredWidthTypeAndValue
            //ExSummary:Retrieves the preferred width type of a table cell.
            Cell firstCell = table.FirstRow.FirstCell;
            PreferredWidthType type = firstCell.CellFormat.PreferredWidth.Type;
            double value = firstCell.CellFormat.PreferredWidth.Value;
            //ExEnd

            Assert.AreEqual(PreferredWidthType.Percent, type);
            Assert.AreEqual(11.16, value);
        }

        [Test]
        public void InsertTableUsingNodeConstructors()
        {
            //ExStart
            //ExFor:Table
            //ExFor:Row
            //ExFor:Row.RowFormat
            //ExFor:RowFormat
            //ExFor:Cell
            //ExFor:Cell.CellFormat
            //ExFor:CellFormat
            //ExFor:CellFormat.Shading
            //ExFor:Cell.FirstParagraph
            //ExId:InsertTableUsingNodeConstructors
            //ExSummary:Shows how to insert a table using the constructors of nodes.
            Document doc = new Document();

            // We start by creating the table object. Note how we must pass the document object
            // to the constructor of each node. This is because every node we create must belong
            // to some document.
            Table table = new Table(doc);
            // Add the table to the document.
            doc.FirstSection.Body.AppendChild(table);

            // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
            // to ensure that the specified node is valid, in this case a valid table should have at least one
            // row and one cell, therefore this method creates them for us.

            // Instead we will handle creating the row and table ourselves. This would be the best way to do this
            // if we were creating a table inside an algorthim for example.
            Row row = new Row(doc);
            row.RowFormat.AllowBreakAcrossPages = true;
            table.AppendChild(row);

            // We can now apply any auto fit settings.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Create a cell and add it to the row
            Cell cell = new Cell(doc);
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            cell.CellFormat.Width = 80;
            
            // Add a paragraph to the cell as well as a new run with some text.
            cell.AppendChild(new Paragraph(doc));
            cell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 1 Text"));

            // Add the cell to the row.
            row.AppendChild(cell);

            // We would then repeat the process for the other cells and rows in the table.
            // We can also speed things up by cloning existing cells and rows.
            row.AppendChild(cell.Clone(false));
            row.LastCell.AppendChild(new Paragraph(doc));
            row.LastCell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 2 Text"));

            doc.Save(MyDir + @"\Artifacts\Table.InsertTableUsingNodes.doc");
            //ExEnd

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Row, true).Count);
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual("Row 1, Cell 1 Text\r\nRow 1, Cell 2 Text", doc.FirstSection.Body.Tables[0].ToString(SaveFormat.Text).Trim());
        }

        //ExStart
        //ExFor:Table
        //ExFor:Row
        //ExFor:Cell
        //ExFor:Table.#ctor(DocumentBase)
        //ExFor:Row.#ctor(DocumentBase)
        //ExFor:Cell.#ctor(DocumentBase)
        //ExId:NestedTableNodeConstructors
        //ExSummary:Shows how to build a nested table without using DocumentBuilder.
        [Test] //ExSkip
        public void NestedTablesUsingNodeConstructors()
        {
            Document doc = new Document();

            // Create the outer table with three rows and four columns.
            Table outerTable = this.CreateTable(doc, 3, 4, "Outer Table");
            // Add it to the document body.
            doc.FirstSection.Body.AppendChild(outerTable);

            // Create another table with two rows and two columns.
            Table innerTable = this.CreateTable(doc, 2, 2, "Inner Table");
            // Add this table to the first cell of the outer table.
            outerTable.FirstRow.FirstCell.AppendChild(innerTable);

            doc.Save(MyDir + @"\Artifacts\Table.CreateNestedTable.doc");

            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count); // ExSkip
            Assert.AreEqual(1, outerTable.FirstRow.FirstCell.Tables.Count); //ExSkip
            Assert.AreEqual(16, outerTable.GetChildNodes(NodeType.Cell, true).Count); //ExSkip
            Assert.AreEqual(4, innerTable.GetChildNodes(NodeType.Cell, true).Count); //ExSkip
        }

        /// <summary>
        /// Creates a new table in the document with the given dimensions and text in each cell.
        /// </summary>
        private Table CreateTable(Document doc, int rowCount, int cellCount, string cellText)
        {
            Table table = new Table(doc);

            // Create the specified number of rows.
            for (int rowId = 1; rowId <= rowCount; rowId++)
            {
                Row row = new Row(doc);
                table.AppendChild(row);

                // Create the specified number of cells for each row.
                for (int cellId = 1; cellId <= cellCount; cellId++)
                {
                    Cell cell = new Cell(doc);
                    row.AppendChild(cell);
                    // Add a blank paragraph to the cell.
                    cell.AppendChild(new Paragraph(doc));

                    // Add the text.
                    cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                }
            }

            return table;
        }
        //ExEnd

        //ExStart
        //ExFor:CellFormat.HorizontalMerge
        //ExFor:CellFormat.VerticalMerge
        //ExFor:CellMerge
        //ExId:CheckCellMerge
        //ExSummary:Prints the horizontal and vertical merge type of a cell.
        [Test] //ExSkip
        public void CheckCellsMerged()
        {
            Document doc = new Document(MyDir + "Table.MergedCells.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Console.WriteLine(this.PrintCellMergeType(cell));
                }
            }

            Assert.AreEqual("The cell at R1, C1 is horizontally merged.", this.PrintCellMergeType(table.FirstRow.FirstCell)); //ExSkip
        }

        public string PrintCellMergeType(Cell cell)
        {
            bool isHorizontallyMerged = cell.CellFormat.HorizontalMerge != CellMerge.None;
            bool isVerticallyMerged = cell.CellFormat.VerticalMerge != CellMerge.None;
            string cellLocation = string.Format("R{0}, C{1}", cell.ParentRow.ParentTable.IndexOf(cell.ParentRow) + 1, cell.ParentRow.IndexOf(cell) + 1);

            if (isHorizontallyMerged && isVerticallyMerged)
                return string.Format("The cell at {0} is both horizontally and vertically merged", cellLocation);
            else if (isHorizontallyMerged)
                return string.Format("The cell at {0} is horizontally merged.", cellLocation);
            else if (isVerticallyMerged)
                return string.Format("The cell at {0} is vertically merged", cellLocation);
            else
                return string.Format("The cell at {0} is not merged", cellLocation);
        }
        //ExEnd

        [Test]
        public void MergeCellRange()
        {
            // Open the document
            Document doc = new Document(MyDir + "Table.Document.doc");

            // Retrieve the first table in the body of the first section.
            Table table = doc.FirstSection.Body.Tables[0];

            //ExStart
            //ExId:MergeCellRange
            //ExSummary:Merges the range of cells between the two specified cells.
            // We want to merge the range of cells found inbetween these two cells.
            Cell cellStartRange = table.Rows[2].Cells[2];
            Cell cellEndRange = table.Rows[3].Cells[3];

            // Merge all the cells between the two specified cells into one.
            MergeCells(cellStartRange, cellEndRange);
            //ExEnd

            // Save the document.
            doc.Save(MyDir + @"\Artifacts\Table.MergeCellRange.doc");

            // Verify the cells were merged
            int mergedCellsCount = 0;
            foreach(Cell cell in table.GetChildNodes(NodeType.Cell, true))
                if(cell.CellFormat.HorizontalMerge != CellMerge.None || cell.CellFormat.HorizontalMerge != CellMerge.None)
                    mergedCellsCount++;

            Assert.AreEqual(4, mergedCellsCount);
            Assert.True(table.Rows[2].Cells[2].CellFormat.HorizontalMerge == CellMerge.First);
            Assert.True(table.Rows[2].Cells[2].CellFormat.VerticalMerge == CellMerge.First);
            Assert.True(table.Rows[3].Cells[3].CellFormat.HorizontalMerge == CellMerge.Previous);
            Assert.True(table.Rows[3].Cells[3].CellFormat.VerticalMerge == CellMerge.Previous);
        }

        //ExStart
        //ExId:MergeCellsMethod
        //ExSummary:A method which merges all cells of a table in the specified range of cells.
        /// <summary>
        /// Merges the range of cells found between the two specified cells both horizontally and vertically. Can span over multiple rows.
        /// </summary>
        public static void MergeCells(Cell startCell, Cell endCell)
        {
            Table parentTable = startCell.ParentRow.ParentTable;

            // Find the row and cell indices for the start and end cell.
            Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell), parentTable.IndexOf(startCell.ParentRow));
            Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow));
            // Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell. 
            Rectangle mergeRange = new Rectangle(Math.Min(startCellPos.X, endCellPos.X), Math.Min(startCellPos.Y, endCellPos.Y), 
                Math.Abs(endCellPos.X - startCellPos.X) + 1, Math.Abs(endCellPos.Y - startCellPos.Y) + 1);

            foreach (Row row in parentTable.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Point currentPos = new Point(row.IndexOf(cell), parentTable.IndexOf(row));

                    // Check if the current cell is inside our merge range then merge it.
                    if (mergeRange.Contains(currentPos))
                    {
                        if (currentPos.X == mergeRange.X)
                            cell.CellFormat.HorizontalMerge = CellMerge.First;
                        else
                            cell.CellFormat.HorizontalMerge = CellMerge.Previous;

                        if (currentPos.Y == mergeRange.Y)
                            cell.CellFormat.VerticalMerge = CellMerge.First;
                        else
                            cell.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                }
            }
        }
        //ExEnd

        [Test]
        public void CombineTables()
        {
            //ExStart
            //ExFor:Table
            //ExFor:Cell.CellFormat
            //ExFor:CellFormat.Borders
            //ExFor:Table.Rows
            //ExFor:Table.FirstRow
            //ExFor:CellFormat.ClearFormatting
            //ExId:CombineTables
            //ExSummary:Shows how to combine the rows from two tables into one.
            // Load the document.
            Document doc = new Document(MyDir + "Table.Document.doc");

            // Get the first and second table in the document.
            // The rows from the second table will be appended to the end of the first table.
            Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);
            Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

            // Append all rows from the current table to the next.
            // Due to the design of tables even tables with different cell count and widths can be joined into one table.
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the empty table container.
            secondTable.Remove();

            doc.Save(MyDir + @"\Artifacts\Table.CombineTables.doc");
            //ExEnd

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(9, doc.FirstSection.Body.Tables[0].Rows.Count);
            Assert.AreEqual(42, doc.FirstSection.Body.Tables[0].GetChildNodes(NodeType.Cell, true).Count);
        }

        [Test]
        public void SplitTable()
        {
            //ExStart
            //ExId:SplitTableAtRow
            //ExSummary:Shows how to split a table into two tables a specific row.
            // Load the document.
            Document doc = new Document(MyDir + "Table.SimpleTable.doc");

            // Get the first table in the document.
            Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);

            // We will split the table at the third row (inclusive).
            Row row = firstTable.Rows[2];

            // Create a new container for the split table.
            Table table = (Table)firstTable.Clone(false);

            // Insert the container after the original.
            firstTable.ParentNode.InsertAfter(table, firstTable);

            // Add a buffer paragraph to ensure the tables stay apart.
            firstTable.ParentNode.InsertAfter(new Paragraph(doc), firstTable);

            Row currentRow;

            do
            {
                currentRow = firstTable.LastRow;
                table.PrependChild(currentRow);
            }
            while (currentRow != row);

            doc.Save(MyDir + @"\Artifacts\Table.SplitTable.doc");
            //ExEnd

            doc = new Document(MyDir + @"\Artifacts\Table.SplitTable.doc");
            // Test we are adding the rows in the correct order and the 
            // selected row was also moved.
            Assert.AreEqual(row, table.FirstRow); 

            Assert.AreEqual(2, firstTable.Rows.Count);
            Assert.AreEqual(2, table.Rows.Count);
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count);
        }
    }
}
