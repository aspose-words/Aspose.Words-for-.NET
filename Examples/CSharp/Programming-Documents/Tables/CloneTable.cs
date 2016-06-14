
using System;
using System.Collections;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class CloneTable
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            CloneCompleteTable(dataDir);
            CloneLastRow(dataDir);                       
        }
        /// <summary>
        /// Shows how to clone complete table.
        /// </summary>
        private static void CloneCompleteTable(string dataDir)
        {
            //ExStart:CloneCompleteTable
            Document doc = new Document(dataDir + "Table.SimpleTable.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Create a clone of the table.
            Table tableClone = (Table)table.Clone(true);

            // Insert the cloned table into the document after the original
            table.ParentNode.InsertAfter(tableClone, table);

            // Insert an empty paragraph between the two tables or else they will be combined into one
            // upon save. This has to do with document validation.
            table.ParentNode.InsertAfter(new Paragraph(doc), table);
            dataDir = dataDir + "Table.CloneTableAndInsert_out_.doc";
           
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:CloneCompleteTable
            Console.WriteLine("\nTable cloned successfully.\nFile saved at " + dataDir);
        }
        /// <summary>
        /// Shows how to clone last row of table.
        /// </summary>
        private static void CloneLastRow(string dataDir)
        {
            //ExStart:CloneLastRow
            Document doc = new Document(dataDir + "Table.SimpleTable.doc");

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

            dataDir = dataDir + "Table.AddCloneRowToTable_out_.doc";
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:CloneLastRow
            Console.WriteLine("\nTable last row cloned successfully.\nFile saved at " + dataDir);
        }
    }
}
