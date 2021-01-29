using Aspose.Words;
using Aspose.Words.Tables;

namespace _03._03_AutofitSettingtoTables
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the document
            Document doc = new Document("../../data/document.doc");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Autofit the first table to the page width.
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            // Save the document to disk.
            doc.Save("AutofitSettingtoTables.docx");
        }
    }
}
