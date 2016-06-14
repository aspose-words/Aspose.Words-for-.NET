
using System;
using System.Collections;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class ApplyFormatting
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            ApplyOutlineBorder(dataDir);
            BuildTableWithBordersEnabled(dataDir);
            ModifyRowFormatting(dataDir);
            ApplyRowFormatting(dataDir);
            ModifyCellFormatting(dataDir);
            FormatTableAndCellWithDifferentBorders(dataDir);
        }
        /// <summary>
        /// Shows how to apply outline border to a table.
        /// </summary>
        private static void ApplyOutlineBorder(string dataDir)
        {
            //ExStart:ApplyOutlineBorder
            Document doc = new Document(dataDir + "Table.EmptyTable.doc");
            
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
            dataDir = dataDir + "Table.SetOutlineBorders_out_.doc";
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:ApplyOutlineBorder
            Console.WriteLine("\nOutline border applied successfully to a table.\nFile saved at " + dataDir);
        }
        /// <summary>
        /// Shows how to build a table with all borders enabled (grid).
        /// </summary>
        private static void BuildTableWithBordersEnabled(string dataDir)
        {
            //ExStart:BuildTableWithBordersEnabled
            Document doc = new Document(dataDir + "Table.EmptyTable.doc");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            // Clear any existing borders from the table.
            table.ClearBorders();
            // Set a green border around and inside the table.
            table.SetBorders(LineStyle.Single, 1.5, Color.Green);

            dataDir = dataDir + "Table.SetAllBorders_out_.doc";
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:BuildTableWithBordersEnabled
            Console.WriteLine("\nTable build successfully with all borders enabled.\nFile saved at " + dataDir);
        }
        /// <summary>
        /// Shows how to modify formatting of a table row.
        /// </summary>
        private static void ModifyRowFormatting(string dataDir)
        {
            //ExStart:ModifyRowFormatting
            Document doc = new Document(dataDir + "Table.Document.doc");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            
            // Retrieve the first row in the table.
            Row firstRow = table.FirstRow;
            // Modify some row level properties.
            firstRow.RowFormat.Borders.LineStyle = LineStyle.None;
            firstRow.RowFormat.HeightRule = HeightRule.Auto;
            firstRow.RowFormat.AllowBreakAcrossPages = true; 
            //ExEnd:ModifyRowFormatting
            Console.WriteLine("\nSome row level properties modified successfully.");
        }
        /// <summary>
        /// Shows how to create a table that contains a single cell and apply row formatting.
        /// </summary>
        private static void ApplyRowFormatting(string dataDir)
        {
            //ExStart:ApplyRowFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the row formatting
            RowFormat rowFormat = builder.RowFormat;
            rowFormat.Height = 100;
            rowFormat.HeightRule = HeightRule.Exactly;
            // These formatting properties are set on the table and are applied to all rows in the table.
            table.LeftPadding = 30;
            table.RightPadding = 30;
            table.TopPadding = 30;
            table.BottomPadding = 30;

            builder.Writeln("I'm a wonderful formatted row.");

            builder.EndRow();
            builder.EndTable();

            dataDir = dataDir + "Table.ApplyRowFormatting_out_.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:ApplyRowFormatting
            Console.WriteLine("\nRow formatting applied successfully.\nFile saved at " + dataDir);
        }
        /// <summary>
        /// Shows how to modify formatting of a table cell.
        /// </summary>
        private static void ModifyCellFormatting(string dataDir)
        {
            //ExStart:ModifyCellFormatting
            Document doc = new Document(dataDir + "Table.Document.doc"); 
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Retrieve the first cell in the table.
            Cell firstCell = table.FirstRow.FirstCell;
            // Modify some cell level properties.
            firstCell.CellFormat.Width = 30; // in points
            firstCell.CellFormat.Orientation = TextOrientation.Downward;
            firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen;
            //ExEnd:ModifyCellFormatting
            Console.WriteLine("\nSome cell level properties modified successfully.");
        }
        /// <summary>
        /// Shows how to format table and cell with different borders and shadings.
        /// </summary>
        private static void FormatTableAndCellWithDifferentBorders(string dataDir)
        {
            //ExStart:FormatTableAndCellWithDifferentBorders
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the borders for the entire table.
            table.SetBorders(LineStyle.Single, 2.0, Color.Black);
            // Set the cell shading for this cell.
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Red;
            builder.Writeln("Cell #1");

            builder.InsertCell();
            // Specify a different cell shading for the second cell.
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Green;
            builder.Writeln("Cell #2");

            // End this row.
            builder.EndRow();

            // Clear the cell formatting from previous operations.
            builder.CellFormat.ClearFormatting();

            // Create the second row.
            builder.InsertCell();

            // Create larger borders for the first cell of this row. This will be different.
            // compared to the borders set for the table.
            builder.CellFormat.Borders.Left.LineWidth = 4.0;
            builder.CellFormat.Borders.Right.LineWidth = 4.0;
            builder.CellFormat.Borders.Top.LineWidth = 4.0;
            builder.CellFormat.Borders.Bottom.LineWidth = 4.0;
            builder.Writeln("Cell #3");

            builder.InsertCell();
            // Clear the cell formatting from the previous cell.
            builder.CellFormat.ClearFormatting();
            builder.Writeln("Cell #4");
            // Save finished document.
            doc.Save(dataDir + "Table.SetBordersAndShading_out_.doc");
            //ExEnd:FormatTableAndCellWithDifferentBorders
            Console.WriteLine("\nformat table and cell with different borders and shadings successfully.\nFile saved at " + dataDir);
        }
    }
}
