using Aspose.Words;
using Aspose.Words.Tables;

namespace Joining_Tables
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"..\..\..\..\Sample Files\";

            // Load a document that contains tables from the local file system.
            Document doc = new Document(filePath + "Tables.docx");

            // Get the first and second table in the document.
            // The rows from the second table will be appended to the end of the first table.
            Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);
            Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

            // Append all rows from the current table to the next.
            // Due to the design of tables even tables with different cell count and widths can be joined into one table.
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the empty table container.
            secondTable.Remove();

            doc.Save(filePath + "Tables Joined.docx");
        }
    }
}
