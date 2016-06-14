
using System;
using System.Collections;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class KeepTablesAndRowsBreaking
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            
            // The below method shows how to disable rows breaking across pages for every row in a table.
            RowFormatDisableBreakAcrossPages(dataDir);
            // The below method shows how to set a table to stay together on the same page.
            KeepTableTogether(dataDir);
           
        }
        public  static void RowFormatDisableBreakAcrossPages(string dataDir)
        {
            //ExStart:RowFormatDisableBreakAcrossPages
            Document doc = new Document(dataDir + "Table.TableAcrossPage.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);            
            // Disable breaking across pages for all rows in the table.
            foreach (Row row in table)
                row.RowFormat.AllowBreakAcrossPages = false;

            dataDir = dataDir + "Table.DisableBreakAcrossPages_out_.doc";
            doc.Save(dataDir);
            //ExEnd:RowFormatDisableBreakAcrossPages
            Console.WriteLine("\nTable rows breaking across pages for every row in a table disabled successfully.\nFile saved at " + dataDir);
        }
        public static void KeepTableTogether(string dataDir)
        {
            //ExStart:KeepTableTogether
           Document doc = new Document(dataDir + "Table.TableAcrossPage.doc");
            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // To keep a table from breaking across a page we need to enable KeepWithNext 
            // for every paragraph in the table except for the last paragraphs in the last 
            // row of the table.
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                // call this method if table's cell is created on the fly
                // newly created cell does not have paragraph inside
                cell.EnsureMinimum();
                foreach (Paragraph para in cell.Paragraphs)
                    if (!(cell.ParentRow.IsLastRow && para.IsEndOfCell))
                        para.ParagraphFormat.KeepWithNext = true;
            }
            dataDir = dataDir + "Table.KeepTableTogether_out_.doc";
            doc.Save(dataDir);
            //ExEnd:KeepTableTogether
            Console.WriteLine("\nTable setup successfully to stay together on the same page.\nFile saved at " + dataDir);            
        }
 
    }
}
