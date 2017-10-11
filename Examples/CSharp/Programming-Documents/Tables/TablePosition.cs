using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class TablePosition
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            GetTablePosition(dataDir);
        }

        private static void GetTablePosition(string dataDir)
        {
            // ExStart:GetTablePosition
            Document doc = new Document(dataDir + "Table.Document.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            if (table.TextWrapping == TextWrapping.Around)
            {
                Console.WriteLine(table.RelativeHorizontalAlignment);
                Console.WriteLine(table.RelativeVerticalAlignment);
            }
            else
            {
                Console.WriteLine(table.Alignment);
            }

            // ExEnd:GetTablePosition
            Console.WriteLine("\nGet the Table position successfully.");
        }
    }
}
