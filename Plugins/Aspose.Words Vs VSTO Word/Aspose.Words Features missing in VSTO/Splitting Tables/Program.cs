using Aspose.Words;
using Aspose.Words.Tables;

namespace Splitting_Tables
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"..\..\..\..\Sample Files\";

            // Load a document that contains tables from the local file system.
            Document doc = new Document(filePath + "Tables.docx");

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

            doc.Save(filePath + "Tables Split.docx");
        }
    }
}
