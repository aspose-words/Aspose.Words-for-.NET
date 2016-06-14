
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Tables;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class FindingIndex
    {
        public static void Run()
        {
            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables() + "Table.SimpleTable.doc";
            Document doc = new Document(dataDir);

            //ExStart:RetrieveTableIndex
            // Get the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
            int tableIndex = allTables.IndexOf(table);            
            //ExEnd:RetrieveTableIndex
            Console.WriteLine("\nTable index is " + tableIndex.ToString());

            //ExStart:RetrieveRowIndex
            int rowIndex = table.IndexOf((Row)table.LastRow);
            //ExEnd:RetrieveRowIndex
            Console.WriteLine("\nRow index is " + rowIndex.ToString());

            Row row = (Row)table.LastRow;
            //ExStart:RetrieveCellIndex
            int cellIndex = row.IndexOf(row.Cells[4]);
            //ExEnd:RetrieveCellIndex
            Console.WriteLine("\nCell index is " + cellIndex.ToString());
        
        }        
    }
}
