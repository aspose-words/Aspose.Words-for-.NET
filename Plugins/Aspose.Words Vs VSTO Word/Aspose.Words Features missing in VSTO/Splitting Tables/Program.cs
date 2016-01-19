using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Splitting_Tables
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"E:\Aspose\Aspose Vs VSTO\Aspose.Words Features missing in VSTO 1.1\Sample Files\";
            
            // Load the document.
            Document doc = new Document(MyDir + "Splitting_Tables.doc");

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

            doc.Save(MyDir + "Splitting_Tables_Out.doc");
        }
    }
}
