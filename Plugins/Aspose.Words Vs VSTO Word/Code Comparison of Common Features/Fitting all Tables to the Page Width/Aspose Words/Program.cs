using Aspose.Words;
using Aspose.Words.Tables;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 1, Row 1, cell 1.");
            builder.InsertCell();
            builder.Write("Table 1, Row 1, cell 2.");
            builder.EndTable();

            builder.InsertParagraph();

            builder.StartTable();
            builder.InsertCell();
            builder.Write("Table 2, Row 1, cell 1.");
            builder.InsertCell();
            builder.Write("Table 2, Row 1, cell 2.");
            builder.EndTable();

            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                table.PreferredWidth = PreferredWidth.FromPercent(100);
            }

            doc.Save("Fitting all Tables to the Page Width.docx");
        }
    }
}
