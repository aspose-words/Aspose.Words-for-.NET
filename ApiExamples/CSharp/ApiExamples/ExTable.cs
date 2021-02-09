// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
            //ExSummary:Shows how to create a table.
            Document doc = new Document();
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Tables contain rows, which contain cells, which may have paragraphs
            // with typical elements such as runs, shapes, and even other tables.
            // Calling the "EnsureMinimum" method on a table will ensure that
            // the table has at least one row, cell, and paragraph.
            Row firstRow = new Row(doc);
            table.AppendChild(firstRow);

            Cell firstCell = new Cell(doc);
            firstRow.AppendChild(firstCell);

            Paragraph paragraph = new Paragraph(doc);
            firstCell.AppendChild(paragraph);

            // Add text to the first call in the first row of the table.
            Run run = new Run(doc, "Hello world!");
            paragraph.AppendChild(run);

            doc.Save(ArtifactsDir + "Table.CreateTable.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.CreateTable.docx");
            table = doc.FirstSection.Body.Tables[0];

            Assert.AreEqual(1, table.Rows.Count);
            Assert.AreEqual(1, table.FirstRow.Cells.Count);
            Assert.AreEqual("Hello world!\a\a", table.GetText().Trim());
        }

        [Test]
        public void Padding()
        {
            //ExStart
            //ExFor:Table.LeftPadding
            //ExFor:Table.RightPadding
            //ExFor:Table.TopPadding
            //ExFor:Table.BottomPadding
            //ExSummary:Shows how to configure content padding in a table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, cell 1.");
            builder.InsertCell();
            builder.Write("Row 1, cell 2.");
            builder.EndTable();
            
            // For every cell in the table, set the distance between its contents and each of its borders. 
            // This table will maintain the minimum padding distance by wrapping text.
            table.LeftPadding = 30;
            table.RightPadding = 60;
            table.TopPadding = 10;
            table.BottomPadding = 90;
            table.PreferredWidth = PreferredWidth.FromPoints(250);

            doc.Save(ArtifactsDir + "DocumentBuilder.SetRowFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.SetRowFormatting.docx");
            table = doc.FirstSection.Body.Tables[0];

            Assert.AreEqual(30.0d, table.LeftPadding);
            Assert.AreEqual(60.0d, table.RightPadding);
            Assert.AreEqual(10.0d, table.TopPadding);
            Assert.AreEqual(90.0d, table.BottomPadding);
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
            //ExSummary:Shows how to modify the format of rows and cells in a table.
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

            // Use the first row's "RowFormat" property to modify the formatting
            // of the contents of all cells in this row.
            RowFormat rowFormat = table.FirstRow.RowFormat;
            rowFormat.Height = 25;
            rowFormat.Borders[BorderType.Bottom].Color = Color.Red;

            // Use the "CellFormat" property of the first cell in the last row to modify the formatting of that cell's contents.
            CellFormat cellFormat = table.LastRow.FirstCell.CellFormat;
            cellFormat.Width = 100;
            cellFormat.Shading.BackgroundPatternColor = Color.Orange;

            doc.Save(ArtifactsDir + "Table.RowCellFormat.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.RowCellFormat.docx");
            table = doc.FirstSection.Body.Tables[0];

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
            //ExSummary:Shows how to iterate through all tables in the document and print the contents of each cell.
            Document doc = new Document(MyDir + "Tables.docx");
            TableCollection tables = doc.FirstSection.Body.Tables;

            Assert.AreEqual(2, tables.ToArray().Length);

            for (int i = 0; i < tables.Count; i++)
            {
                Console.WriteLine($"Start of Table {i}");

                RowCollection rows = tables[i].Rows;

                // We can use the "ToArray" method on a row collection to clone it into an array.
                Assert.AreEqual(rows, rows.ToArray());
                Assert.AreNotSame(rows, rows.ToArray());

                for (int j = 0; j < rows.Count; j++)
                {
                    Console.WriteLine($"\tStart of Row {j}");

                    CellCollection cells = rows[j].Cells;

                    // We can use the "ToArray" method on a cell collection to clone it into an array.
                    Assert.AreEqual(cells, cells.ToArray());
                    Assert.AreNotSame(cells, cells.ToArray());

                    for (int k = 0; k < cells.Count; k++)
                    {
                        string cellText = cells[k].ToString(SaveFormat.Text).Trim();
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
        //ExSummary:Shows how to find out if a tables are nested.
        [Test] //ExSkip
        public void CalculateDepthOfNestedTables()
        {
            Document doc = new Document(MyDir + "Nested tables.docx");
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            Assert.AreEqual(5, tables.Count); //ExSkip

            for (int i = 0; i < tables.Count; i++)
            {
                Table table = (Table)tables[i];

                // Find out if any cells in the table have other tables as children.
                int count = GetChildTableCount(table);
                Console.WriteLine("Table #{0} has {1} tables directly within its cells", i, count);

                // Find out if the table is nested inside another table, and, if so, at what depth.
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
        /// </summary>
        /// <returns>
        /// An integer indicating the nesting depth of the table (number of parent table nodes).
        /// </returns>
        private static int GetNestedDepthOfTable(Table table)
        {
            int depth = 0;
            Node parent = table.GetAncestor(table.NodeType);

            while (parent != null)
            {
                depth++;
                parent = parent.GetAncestor(typeof(Table));
            }

            return depth;
        }

        /// <summary>
        /// Determines if a table contains any immediate child table within its cells.
        /// Do not recursively traverse through those tables to check for further tables.
        /// </summary>
        /// <returns>
        /// Returns true if at least one child cell contains a table.
        /// Returns false if no cells in the table contain a table.
        /// </returns>
        private static int GetChildTableCount(Table table)
        {
            int childTableCount = 0;

            foreach (Row row in table.Rows.OfType<Row>())
            {
                foreach (Cell Cell in row.Cells.OfType<Cell>())
                {
                    TableCollection childTables = Cell.Tables;

                    if (childTables.Count > 0)
                        childTableCount++;
                }
            }

            return childTableCount;
        }
        //ExEnd

        [Test]
        public void EnsureTableMinimum()
        {
            //ExStart
            //ExFor:Table.EnsureMinimum
            //ExSummary:Shows how to ensure that a table node contains the nodes we need to add content.
            Document doc = new Document();
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            // Tables contain rows, which contain cells, which may contain paragraphs
            // with typical elements such as runs, shapes, and even other tables.
            // Our new table has none of these nodes, and we cannot add contents to it until it does.
            Assert.AreEqual(0, table.GetChildNodes(NodeType.Any, true).Count);

            // Calling the "EnsureMinimum" method on a table will ensure that
            // the table has at least one row and one cell with an empty paragraph.
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
            //ExSummary:Shows how to ensure a row node contains the nodes we need to begin adding content to it.
            Document doc = new Document();
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);
            Row row = new Row(doc);
            table.AppendChild(row);

            // Rows contain cells, containing paragraphs with typical elements such as runs, shapes, and even other tables.
            // Our new row has none of these nodes, and we cannot add contents to it until it does.
            Assert.AreEqual(0, row.GetChildNodes(NodeType.Any, true).Count);

            // Calling the "EnsureMinimum" method on a table will ensure that
            // the table has at least one cell with an empty paragraph.
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
            //ExSummary:Shows how to ensure a cell node contains the nodes we need to begin adding content to it.
            Document doc = new Document();
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);
            Row row = new Row(doc);
            table.AppendChild(row);
            Cell cell = new Cell(doc);
            row.AppendChild(cell);

            // Cells may contain paragraphs with typical elements such as runs, shapes, and even other tables.
            // Our new cell does not have any paragraphs, and we cannot add contents such as run and shape nodes to it until it does.
            Assert.AreEqual(0, cell.GetChildNodes(NodeType.Any, true).Count);

            // Calling the "EnsureMinimum" method on a cell will ensure that
            // the cell has at least one empty paragraph, which we can then add contents to.
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
            //ExSummary:Shows how to apply an outline border to a table.
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = doc.FirstSection.Body.Tables[0];

            // Align the table to the center of the page.
            table.Alignment = TableAlignment.Center;

            // Clear any existing borders and shading from the table.
            table.ClearBorders();
            table.ClearShading();

            // Add green borders to the outline of the table.
            table.SetBorder(BorderType.Left, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Right, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Green, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Green, true);

            // Fill the cells with a light green solid color.
            table.SetShading(TextureIndex.TextureSolid, Color.LightGreen, Color.Empty);

            doc.Save(ArtifactsDir + "Table.SetOutlineBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.SetOutlineBorders.docx");
            table = doc.FirstSection.Body.Tables[0];

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
        public void SetBorders()
        {
            //ExStart
            //ExFor:Table.SetBorders
            //ExSummary:Shows how to format of all of a table's borders at once.
            Document doc = new Document(MyDir + "Tables.docx");
            Table table = doc.FirstSection.Body.Tables[0];

            // Clear all existing borders from the table.
            table.ClearBorders();

            // Set a single green line to serve as every outer and inner border of this table.
            table.SetBorders(LineStyle.Single, 1.5, Color.Green);

            doc.Save(ArtifactsDir + "Table.SetBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.SetBorders.docx");
            table = doc.FirstSection.Body.Tables[0];
            
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
            Table table = doc.FirstSection.Body.Tables[0];

            // Use the first row's "RowFormat" property to set formatting that modifies that entire row's appearance.
            Row firstRow = table.FirstRow;
            firstRow.RowFormat.Borders.LineStyle = LineStyle.None;
            firstRow.RowFormat.HeightRule = HeightRule.Auto;
            firstRow.RowFormat.AllowBreakAcrossPages = true;

            doc.Save(ArtifactsDir + "Table.RowFormat.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.RowFormat.docx");
            table = doc.FirstSection.Body.Tables[0];

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
            Table table = doc.FirstSection.Body.Tables[0];
            Cell firstCell = table.FirstRow.FirstCell;

            // Use a cell's "CellFormat" property to set formatting that modifies the appearance of that cell.
            firstCell.CellFormat.Width = 30;
            firstCell.CellFormat.Orientation = TextOrientation.Downward;
            firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen;

            doc.Save(ArtifactsDir + "Table.CellFormat.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.CellFormat.docx");

            table = doc.FirstSection.Body.Tables[0];
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

            Table table = doc.FirstSection.Body.Tables[0];

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

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Hello world!");
            builder.EndTable();

            // Modify the color and thickness of the top border.
            Border topBorder = table.FirstRow.RowFormat.Borders[BorderType.Top];
            table.SetBorder(BorderType.Top, LineStyle.Double, 1.5, Color.Red, true);

            Assert.AreEqual(1.5d, topBorder.LineWidth);
            Assert.AreEqual(Color.Red.ToArgb(), topBorder.Color.ToArgb());
            Assert.AreEqual(LineStyle.Double, topBorder.LineStyle);

            // Clear the borders of all cells in the table, and then save the document.
            table.ClearBorders();
            Assert.Throws<AssertionException>(() => Assert.AreEqual(Color.Empty.ToArgb(), topBorder.Color.ToArgb())); //ExSkip
            doc.Save(ArtifactsDir + "Table.ClearBorders.docx");

            // Verify the values of the table's properties after re-opening the document.
            doc = new Document(ArtifactsDir + "Table.ClearBorders.docx");
            table = doc.FirstSection.Body.Tables[0];
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

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Carrots");
            builder.InsertCell();
            builder.Write("50");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Potatoes");
            builder.InsertCell();
            builder.Write("50");
            builder.EndTable();

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = true;
            options.FindWholeWordsOnly = true;

            // Perform a find-and-replace operation on an entire table.
            table.Range.Replace("Carrots", "Eggs", options);

            // Perform a find-and-replace operation on the last cell of the last row of the table.
            table.LastRow.LastCell.Range.Replace("50", "20", options);

            Assert.AreEqual("Eggs\a50\a\a" +
                            "Potatoes\a20\a\a", table.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void PrintTableRange()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = doc.FirstSection.Body.Tables[0];

            // The range text will include control characters such as "\a" for a cell.
            // You can call ToString on the desired node to retrieve the plain text content.

            // Print the plain text range of the table to the screen.
            Console.WriteLine("Contents of the table: ");
            Console.WriteLine(table.Range.Text);
            
            // Print the contents of the second row to the screen.
            Console.WriteLine("\nContents of the row: ");
            Console.WriteLine(table.Rows[1].Range.Text);

            // Print the contents of the last cell in the table to the screen.
            Console.WriteLine("\nContents of the cell: ");
            Console.WriteLine(table.LastRow.LastCell.Range.Text);
            
            Assert.AreEqual("\aColumn 1\aColumn 2\aColumn 3\aColumn 4\a\a", table.Rows[1].Range.Text);
            Assert.AreEqual("Cell 12 contents\a", table.LastRow.LastCell.Range.Text);
        }

        [Test]
        public void CloneTable()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = doc.FirstSection.Body.Tables[0];

            Table tableClone = (Table) table.Clone(true);

            // Insert the cloned table into the document after the original.
            table.ParentNode.InsertAfter(tableClone, table);

            // Insert an empty paragraph between the two tables.
            table.ParentNode.InsertAfter(new Paragraph(doc), table);

            doc.Save(ArtifactsDir + "Table.CloneTable.doc");
            
            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(table.Range.Text, tableClone.Range.Text);

            foreach (Cell cell in tableClone.GetChildNodes(NodeType.Cell, true).OfType<Cell>())
                cell.RemoveAllChildren();
            
            Assert.AreEqual(string.Empty, tableClone.ToString(SaveFormat.Text).Trim());
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AllowBreakAcrossPages(bool allowBreakAcrossPages)
        {
            //ExStart
            //ExFor:RowFormat.AllowBreakAcrossPages
            //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
            Document doc = new Document(MyDir + "Table spanning two pages.docx");
            Table table = doc.FirstSection.Body.Tables[0];

            // Set the "AllowBreakAcrossPages" property to "false" to keep the row
            // in one piece if a table spans two pages, which break up along that row.
            // If the row is too big to fit in one page, Microsoft Word will push it down to the next page.
            // Set the "AllowBreakAcrossPages" property to "true" to allow the row to break up across two pages.
            foreach (Row row in table.OfType<Row>())
                row.RowFormat.AllowBreakAcrossPages = allowBreakAcrossPages;

            doc.Save(ArtifactsDir + "Table.AllowBreakAcrossPages.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.AllowBreakAcrossPages.docx");
            table = doc.FirstSection.Body.Tables[0];

            Assert.AreEqual(3, table.Rows.Count(r => ((Row)r).RowFormat.AllowBreakAcrossPages == allowBreakAcrossPages));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AllowAutoFitOnTable(bool allowAutoFit)
        {
            //ExStart
            //ExFor:Table.AllowAutoFit
            //ExSummary:Shows how to enable/disable automatic table cell resizing.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                          "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                          "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.EndRow();
            builder.EndTable();

            // Set the "AllowAutoFit" property to "false" to get the table to maintain the dimensions
            // of all its rows and cells, and truncate contents if they get too large to fit.
            // Set the "AllowAutoFit" property to "true" to allow the table to change its cells' width and height
            // to accommodate their contents.
            table.AllowAutoFit = allowAutoFit;

            doc.Save(ArtifactsDir + "Table.AllowAutoFitOnTable.html");
            //ExEnd

            if (allowAutoFit)
            {
                TestUtil.FileContainsString(
                    "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
                    ArtifactsDir + "Table.AllowAutoFitOnTable.html");
                TestUtil.FileContainsString(
                    "<td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
                    ArtifactsDir + "Table.AllowAutoFitOnTable.html");
            }
            else
            {
                TestUtil.FileContainsString(
                    "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
                    ArtifactsDir + "Table.AllowAutoFitOnTable.html");
                TestUtil.FileContainsString(
                    "<td style=\"width:7.2pt; border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
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
            Table table = doc.FirstSection.Body.Tables[0];

            // Enabling KeepWithNext for every paragraph in the table except for the
            // last ones in the last row will prevent the table from splitting across multiple pages.
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
            table = doc.FirstSection.Body.Tables[0];

            foreach (Paragraph para in table.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
                if (para.IsEndOfCell && ((Cell)para.ParentNode).ParentRow.IsLastRow)
                    Assert.False(para.ParagraphFormat.KeepWithNext);
                else
                    Assert.True(para.ParagraphFormat.KeepWithNext);
        }

        [Test]
        public void GetIndexOfTableElements()
        {
            //ExStart
            //ExFor:NodeCollection.IndexOf(Node)
            //ExSummary:Shows how to get the index of a node in a collection.
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = doc.FirstSection.Body.Tables[0];
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
            //ExSummary:Shows how to verify the preferred width type and value of a table cell.
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = doc.FirstSection.Body.Tables[0];
            Cell firstCell = table.FirstRow.FirstCell;

            Assert.AreEqual(PreferredWidthType.Percent, firstCell.CellFormat.PreferredWidth.Type);
            Assert.AreEqual(11.16d, firstCell.CellFormat.PreferredWidth.Value);
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

            table.CellSpacing = 3;

            // Set the "AllowCellSpacing" property to "true" to enable spacing between cells
            // with a magnitude equal to the value of the "CellSpacing" property, in points.
            // Set the "AllowCellSpacing" property to "false" to disable cell spacing
            // and ignore the value of the "CellSpacing" property.
            table.AllowCellSpacing = allowCellSpacing;

            doc.Save(ArtifactsDir + "Table.AllowCellSpacing.html");

            // Adjusting the "CellSpacing" property will automatically enable cell spacing.
            table.CellSpacing = 5;

            Assert.True(table.AllowCellSpacing);
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.AllowCellSpacing.html");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(allowCellSpacing, table.AllowCellSpacing);

            if (allowCellSpacing)
                Assert.AreEqual(3.0d, table.CellSpacing);
            else
                Assert.AreEqual(0.0d, table.CellSpacing);

            TestUtil.FileContainsString(
                allowCellSpacing
                    ? "<td style=\"border-style:solid; border-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border:0.5pt single\">"
                    : "<td style=\"border-right-style:solid; border-right-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; " +
                      "padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-bottom:0.5pt single; -aw-border-right:0.5pt single\">",
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
        //ExSummary:Shows how to build a nested table without using a document builder.
        [Test] //ExSkip
        public void CreateNestedTable()
        {
            Document doc = new Document();

            // Create the outer table with three rows and four columns, and then add it to the document.
            Table outerTable = CreateTable(doc, 3, 4, "Outer Table");
            doc.FirstSection.Body.AppendChild(outerTable);

            // Create another table with two rows and two columns and then insert it into the first table's first cell.
            Table innerTable = CreateTable(doc, 2, 2, "Inner Table");
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

            for (int rowId = 1; rowId <= rowCount; rowId++)
            {
                Row row = new Row(doc);
                table.AppendChild(row);

                for (int cellId = 1; cellId <= cellCount; cellId++)
                {
                    Cell cell = new Cell(doc);
                    cell.AppendChild(new Paragraph(doc));
                    cell.FirstParagraph.AppendChild(new Run(doc, cellText));

                    row.AppendChild(cell);
                }
            }

            // You can use the "Title" and "Description" properties to add a title and description respectively to your table.
            // The table must have at least one row before we can use these properties.
            // These properties are meaningful for ISO / IEC 29500 compliant .docx documents (see the OoxmlCompliance class).
            // If we save the document to pre-ISO/IEC 29500 formats, Microsoft Word ignores these properties.
            table.Title = "Aspose table title";
            table.Description = "Aspose table description";

            return table;
        }
        //ExEnd

        private void TestCreateNestedTable(Document doc)
        {
            Table outerTable = doc.FirstSection.Body.Tables[0];
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
            Table table = doc.FirstSection.Body.Tables[0];

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
            Document doc = new Document(MyDir + "Tables.docx");

            Table table = doc.FirstSection.Body.Tables[0];

            // We want to merge the range of cells found in between these two cells.
            Cell cellStartRange = table.Rows[2].Cells[2];
            Cell cellEndRange = table.Rows[3].Cells[3];

            // Merge all the cells between the two specified cells into one.
            MergeCells(cellStartRange, cellEndRange);

            doc.Save(ArtifactsDir + "Table.MergeCellRange.doc");

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
        /// Merges the range of cells found between the two specified cells both horizontally and vertically.
        /// Can span over multiple rows.
        /// </summary>
        public static void MergeCells(Cell startCell, Cell endCell)
        {
            Table parentTable = startCell.ParentRow.ParentTable;

            // Find the row and cell indices for the start and end cells.
            Point startCellPos = new Point(startCell.ParentRow.IndexOf(startCell),
                parentTable.IndexOf(startCell.ParentRow));
            Point endCellPos = new Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow));

            // Create a range of cells to be merged based on these indices.
            // Inverse each index if the end cell is before the start cell.
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

                    // Check if the current cell is inside our merge range, then merge it.
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
            Document doc = new Document(MyDir + "Tables.docx");

            // Below are two ways of getting a table from a document.
            // 1 -  From the "Tables" collection of a Body node:
            Table firstTable = doc.FirstSection.Body.Tables[0];

            // 2 -  Using the "GetChild" method:
            Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

            // Append all rows from the current table to the next.
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the empty table container.
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
            Document doc = new Document(MyDir + "Tables.docx");

            Table firstTable = doc.FirstSection.Body.Tables[0];

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

            doc = DocumentHelper.SaveOpen(doc);

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

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);

            builder.Font.Size = 16;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Set the "TextWrapping" property to "TextWrapping.Around" to get the table to wrap text around it,
            // and push it down into the paragraph below by setting the position.
            table.TextWrapping = TextWrapping.Around;
            table.AbsoluteHorizontalDistance = 100;
            table.AbsoluteVerticalDistance = 20;

            doc.Save(ArtifactsDir + "Table.WrapText.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.WrapText.docx");
            table = doc.FirstSection.Body.Tables[0];

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
            //ExSummary:Shows how to work with floating tables properties.
            Document doc = new Document(MyDir + "Table wrapped by text.docx");

            Table table = doc.FirstSection.Body.Tables[0];

            if (table.TextWrapping == TextWrapping.Around)
            {
                Assert.AreEqual(RelativeHorizontalPosition.Margin, table.HorizontalAnchor);
                Assert.AreEqual(RelativeVerticalPosition.Paragraph, table.VerticalAnchor);
                Assert.AreEqual(false, table.AllowOverlap);

                // Only Margin, Page, Column available in RelativeHorizontalPosition for HorizontalAnchor setter.
                // The ArgumentException will be thrown for any other values.
                table.HorizontalAnchor = RelativeHorizontalPosition.Column;

                // Only Margin, Page, Paragraph available in RelativeVerticalPosition for VerticalAnchor setter.
                // The ArgumentException will be thrown for any other values.
                table.VerticalAnchor = RelativeVerticalPosition.Page;
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
            //ExSummary:Shows how set the location of floating tables.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 1, cell 1");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);

            // Set the table's location to a place on the page, such as, in this case, the bottom right corner.
            table.RelativeVerticalAlignment = VerticalAlignment.Bottom;
            table.RelativeHorizontalAlignment = HorizontalAlignment.Right;

            table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 2, cell 1");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);

            // We can also set a horizontal and vertical offset in points from the paragraph's location where we inserted the table. 
            table.AbsoluteVerticalDistance = 50;
            table.AbsoluteHorizontalDistance = 100;

            doc.Save(ArtifactsDir + "Table.ChangeFloatingTableProperties.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.ChangeFloatingTableProperties.docx");
            table = doc.FirstSection.Body.Tables[0];

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

            // Setting the style properties of a table may affect the properties of the table itself.
            Assert.True(table.Bidi);
            Assert.AreEqual(5.0d, table.CellSpacing);
            Assert.AreEqual("MyTableStyle1", table.StyleName);

            doc.Save(ArtifactsDir + "Table.TableStyleCreation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.TableStyleCreation.docx");
            table = doc.FirstSection.Body.Tables[0];

            Assert.True(table.Bidi);
            Assert.AreEqual(5.0d, table.CellSpacing);
            Assert.AreEqual("MyTableStyle1", table.StyleName);
            Assert.AreEqual(20.0d, tableStyle.BottomPadding);
            Assert.AreEqual(5.0d, tableStyle.LeftPadding);
            Assert.AreEqual(10.0d, tableStyle.RightPadding);
            Assert.AreEqual(20.0d, tableStyle.TopPadding);
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
            //ExSummary:Shows how to set the position of a table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two ways of aligning a table horizontally.
            // 1 -  Use the "Alignment" property to align it to a location on the page, such as the center:
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.Alignment = TableAlignment.Center;
            tableStyle.Borders.Color = Color.Blue;
            tableStyle.Borders.LineStyle = LineStyle.Single;

            // Insert a table and apply the style we created to it.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Aligned to the center of the page");
            builder.EndTable();
            table.PreferredWidth = PreferredWidth.FromPoints(300);
            
            table.Style = tableStyle;

            // 2 -  Use the "LeftIndent" to specify an indent from the left margin of the page:
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

            doc.Save(ArtifactsDir + "Table.SetTableAlignment.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.SetTableAlignment.docx");

            tableStyle = (TableStyle)doc.Styles["MyTableStyle1"];

            Assert.AreEqual(TableAlignment.Center, tableStyle.Alignment);
            Assert.AreEqual(tableStyle, doc.FirstSection.Body.Tables[0].Style);

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

            // Create a custom table style.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");

            // Conditional styles are formatting changes that affect only some of the table's cells
            // based on a predicate, such as the cells being in the last row.
            // Below are three ways of accessing a table style's conditional styles from the "ConditionalStyles" collection.
            // 1 -  By style type:
            tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.AliceBlue;

            // 2 -  By index:
            tableStyle.ConditionalStyles[0].Borders.Color = Color.Black;
            tableStyle.ConditionalStyles[0].Borders.LineStyle = LineStyle.DotDash;
            Assert.AreEqual(ConditionalStyleType.FirstRow, tableStyle.ConditionalStyles[0].Type);

            // 3 -  As a property:
            tableStyle.ConditionalStyles.FirstRow.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Apply padding and text formatting to conditional styles.
            tableStyle.ConditionalStyles.LastRow.BottomPadding = 10;
            tableStyle.ConditionalStyles.LastRow.LeftPadding = 10;
            tableStyle.ConditionalStyles.LastRow.RightPadding = 10;
            tableStyle.ConditionalStyles.LastRow.TopPadding = 10;
            tableStyle.ConditionalStyles.LastColumn.Font.Bold = true;

            // List all possible style conditions.
            using (IEnumerator<ConditionalStyle> enumerator = tableStyle.ConditionalStyles.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    ConditionalStyle currentStyle = enumerator.Current;
                    if (currentStyle != null) Console.WriteLine(currentStyle.Type);
                }
            }

            // Apply the custom style, which contains all conditional styles, to the table.
            table.Style = tableStyle;

            // Our style applies some conditional styles by default.
            Assert.AreEqual(TableStyleOptions.FirstRow | TableStyleOptions.FirstColumn | TableStyleOptions.RowBands, 
                table.StyleOptions);

            // We will need to enable all other styles ourselves via the "StyleOptions" property.
            table.StyleOptions = table.StyleOptions | TableStyleOptions.LastRow | TableStyleOptions.LastColumn;

            doc.Save(ArtifactsDir + "Table.ConditionalStyles.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.ConditionalStyles.docx");
            table = doc.FirstSection.Body.Tables[0];

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

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("First row");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Last row");
            builder.EndTable();

            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            table.Style = tableStyle;

            // Set the table style to color the borders of the first row of the table in red.
            tableStyle.ConditionalStyles.FirstRow.Borders.Color = Color.Red;

            // Set the table style to color the borders of the last row of the table in blue.
            tableStyle.ConditionalStyles.LastRow.Borders.Color = Color.Blue;

            // Below are two ways of using the "ClearFormatting" method to clear the conditional styles.
            // 1 -  Clear the conditional styles for a specific part of a table:
            tableStyle.ConditionalStyles[0].ClearFormatting();

            Assert.AreEqual(Color.Empty, tableStyle.ConditionalStyles.FirstRow.Borders.Color);

            // 2 -  Clear the conditional styles for the entire table:
            tableStyle.ConditionalStyles.ClearFormatting();

            Assert.True(tableStyle.ConditionalStyles.All(s => s.Borders.Color == Color.Empty));
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

            // We can configure a conditional style of a table to apply a different color to the row/column,
            // based on whether the row/column is even or odd, creating an alternating color pattern.
            // We can also apply a number n to the row/column banding,
            // meaning that the color alternates after every n rows/columns instead of one.
            // Create a table where single columns and rows will band the columns will banded in threes.
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

            // Apply a line style to all the borders of the table.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.Borders.Color = Color.Black;
            tableStyle.Borders.LineStyle = LineStyle.Double;

            // Set the two colors, which will alternate over every 3 rows.
            tableStyle.RowStripe = 3;
            tableStyle.ConditionalStyles[ConditionalStyleType.OddRowBanding].Shading.BackgroundPatternColor = Color.LightBlue;
            tableStyle.ConditionalStyles[ConditionalStyleType.EvenRowBanding].Shading.BackgroundPatternColor = Color.LightCyan;

            // Set a color to apply to every even column, which will override any custom row coloring.
            tableStyle.ColumnStripe = 1;
            tableStyle.ConditionalStyles[ConditionalStyleType.EvenColumnBanding].Shading.BackgroundPatternColor = Color.LightSalmon;

            table.Style = tableStyle;

            // The "StyleOptions" property enables row banding by default.
            Assert.AreEqual(TableStyleOptions.FirstRow | TableStyleOptions.FirstColumn | TableStyleOptions.RowBands,
                table.StyleOptions);

            // Use the "StyleOptions" property also to enable column banding.
            table.StyleOptions = table.StyleOptions | TableStyleOptions.ColumnBands;

            doc.Save(ArtifactsDir + "Table.AlternatingRowStyles.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Table.AlternatingRowStyles.docx");
            table = doc.FirstSection.Body.Tables[0];
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

            // Microsoft Word does not write merge flags anymore, defining merged cells by width instead.
            // Aspose.Words by default define only 5 cells in a row, and none of them have the horizontal merge flag,
            // even though there were 7 cells in the row before the horizontal merging took place.
            Table table = doc.FirstSection.Body.Tables[0];
            Row row = table.Rows[0];

            Assert.AreEqual(5, row.Cells.Count);
            Assert.True(row.Cells.All(c => ((Cell)c).CellFormat.HorizontalMerge == CellMerge.None));

            // Use the "ConvertToHorizontallyMergedCells" method to convert cells horizontally merged
            // by its width to the cell horizontally merged by flags.
            // Now, we have 7 cells, and some of them have horizontal merge values.
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