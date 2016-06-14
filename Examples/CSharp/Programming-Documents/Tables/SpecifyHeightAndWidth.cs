
using System;
using System.Collections;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class SpecifyHeightAndWidth
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            AutoFitToPageWidth(dataDir);
            SetPreferredWidthSettings(dataDir);
            RetrievePreferredWidthType(dataDir);    
        }
        /// <summary>
        /// Shows how to set a table to auto fit to 50% of the page width.
        /// </summary>
        private static void AutoFitToPageWidth(string dataDir)
        {
            //ExStart:AutoFitToPageWidth
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with a width that takes up half the page width.
            Table table = builder.StartTable();

            // Insert a few cells
            builder.InsertCell();
            table.PreferredWidth = PreferredWidth.FromPercent(50);
            builder.Writeln("Cell #1");

            builder.InsertCell();
            builder.Writeln("Cell #2");

            builder.InsertCell();
            builder.Writeln("Cell #3");

            dataDir = dataDir + "Table.PreferredWidth_out_.doc";
           
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:AutoFitToPageWidth
            Console.WriteLine("\nTable autofit successfully to 50% of the page width.\nFile saved at " + dataDir);
        }
        /// <summary>
        /// Shows how to set the different preferred width settings.
        /// </summary>
        private static void SetPreferredWidthSettings(string dataDir)
        {
            //ExStart:SetPreferredWidthSettings
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table row made up of three cells which have different preferred widths.
            Table table = builder.StartTable();

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
            builder.Writeln("Cell automatically sized. The size of this cell is calculated from the table preferred width.");
            builder.Writeln("In this case the cell will fill up the rest of the available space.");

            dataDir = dataDir + "Table.CellPreferredWidths_out_.doc";
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:SetPreferredWidthSettings
            Console.WriteLine("\nDifferent preferred width settings set successfully.\nFile saved at " + dataDir);
        }
        /// <summary>
        /// Shows how to retrieves the preferred width type of a table cell.
        /// </summary>
        private static void RetrievePreferredWidthType(string dataDir)
        {
            //ExStart:RetrievePreferredWidthType
            Document doc = new Document(dataDir + "Table.SimpleTable.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            //ExStart:AllowAutoFit
            table.AllowAutoFit = true;
            //ExEnd:AllowAutoFit

            Cell firstCell = table.FirstRow.FirstCell;
            PreferredWidthType type = firstCell.CellFormat.PreferredWidth.Type;
            double value = firstCell.CellFormat.PreferredWidth.Value;

            //ExEnd:RetrievePreferredWidthType
            Console.WriteLine("\nTable preferred width type value is " + value.ToString());
        }
    }
}
