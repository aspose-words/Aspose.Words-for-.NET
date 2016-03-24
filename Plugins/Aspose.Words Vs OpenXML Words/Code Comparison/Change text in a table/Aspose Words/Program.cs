using Aspose.Words;
using Aspose.Words.Tables;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("Change text in a table.doc");

            // Get the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
             
            // Replace any instances of our string in the last cell of the table only.
            table.Rows[1].Cells[2].Range.Replace("Mr", "test", true, true);
            doc.Save("Change text in a table.doc");
        }
    }
}
