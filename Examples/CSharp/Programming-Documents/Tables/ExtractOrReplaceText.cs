
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ExtractText
    {
        public static void Run()
        {
           
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables() + "Table.SimpleTable.doc";
            ExtractPrintText(dataDir);
            ReplaceText(dataDir);
        
        }
        private static void ExtractPrintText(string dataDir)
        {
            //ExStart:ExtractText
            Document doc = new Document(dataDir);

            // Get the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // The range text will include control characters such as "\a" for a cell.
            // You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.

            // Print the plain text range of the table to the screen.
            Console.WriteLine("Contents of the table: ");
            Console.WriteLine(table.Range.Text);
            //ExEnd:ExtractText   

            //ExStart:PrintTextRangeOFRowAndTable
            // Print the contents of the second row to the screen.
            Console.WriteLine("\nContents of the row: ");
            Console.WriteLine(table.Rows[1].Range.Text);

            // Print the contents of the last cell in the table to the screen.
            Console.WriteLine("\nContents of the cell: ");
            Console.WriteLine(table.LastRow.LastCell.Range.Text);
            //ExEnd:PrintTextRangeOFRowAndTable
        }
        private static void ReplaceText(string dataDir)
        {
            //ExStart:ReplaceText
            Document doc = new Document(dataDir);

            // Get the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Replace any instances of our string in the entire table.
            table.Range.Replace("Carrots", "Eggs", new FindReplaceOptions(FindReplaceDirection.Forward));
            // Replace any instances of our string in the last cell of the table only.
            table.LastRow.LastCell.Range.Replace("50", "20", new FindReplaceOptions(FindReplaceDirection.Forward));

            dataDir = RunExamples.GetDataDir_WorkingWithTables() + "Table.ReplaceCellText_out_.doc";
            doc.Save(dataDir); 
            //ExEnd:ReplaceText    
            Console.WriteLine("\nText replaced successfully.\nFile saved at " + dataDir);
        }
    }
}
