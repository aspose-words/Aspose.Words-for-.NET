
using System;
using System.Collections;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class ApplyStyle
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            BuildTableWithStyle(dataDir);
            ExpandFormattingOnCellsAndRowFromStyle(dataDir);                       
        }
        /// <summary>
        /// Shows how to build a new table with a table style applied.
        /// </summary>
        private static void BuildTableWithStyle(string dataDir)
        {
            //ExStart:BuildTableWithStyle
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            // We must insert at least one row first before setting any table formatting.
            builder.InsertCell();
            // Set the table style used based of the unique style identifier.
            // Note that not all table styles are available when saving as .doc format.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
            // Apply which features should be formatted by the style.
            table.StyleOptions = TableStyleOptions.FirstColumn | TableStyleOptions.RowBands | TableStyleOptions.FirstRow;
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Continue with building the table as normal.
            builder.Writeln("Item");
            builder.CellFormat.RightPadding = 40;
            builder.InsertCell();
            builder.Writeln("Quantity (kg)");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Apples");
            builder.InsertCell();
            builder.Writeln("20");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Bananas");
            builder.InsertCell();
            builder.Writeln("40");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Carrots");
            builder.InsertCell();
            builder.Writeln("50");
            builder.EndRow();

            dataDir = dataDir + "DocumentBuilder.SetTableStyle_out_.docx";
           
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:BuildTableWithStyle
            Console.WriteLine("\nTable created successfully with table style.\nFile saved at " + dataDir);
        }
        /// <summary>
        /// Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
        /// </summary>
        private static void ExpandFormattingOnCellsAndRowFromStyle(string dataDir)
        {
            //ExStart:ExpandFormattingOnCellsAndRowFromStyle
            Document doc = new Document(dataDir + "Table.TableStyle.docx");

            // Get the first cell of the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Cell firstCell = table.FirstRow.FirstCell;

            // First print the color of the cell shading. This should be empty as the current shading
            // is stored in the table style.
            Color cellShadingBefore = firstCell.CellFormat.Shading.BackgroundPatternColor;
            Console.WriteLine("Cell shading before style expansion: " + cellShadingBefore.ToString());

            // Expand table style formatting to direct formatting.
            doc.ExpandTableStylesToDirectFormatting();

            // Now print the cell shading after expanding table styles. A blue background pattern color
            // should have been applied from the table style.
            Color cellShadingAfter = firstCell.CellFormat.Shading.BackgroundPatternColor;
            Console.WriteLine("Cell shading after style expansion: " + cellShadingAfter.ToString());
            //ExEnd:ExpandFormattingOnCellsAndRowFromStyle
            
        }
    }
}
