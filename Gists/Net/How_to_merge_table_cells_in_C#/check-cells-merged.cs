// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Table with merged cells.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

foreach (Row row in table.Rows)
{
    foreach (Cell cell in row.Cells)
    {
        Console.WriteLine(PrintCellMergeType(cell));
    }
}
