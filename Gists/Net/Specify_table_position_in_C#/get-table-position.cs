// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

if (table.TextWrapping == TextWrapping.Around)
{
    Console.WriteLine(table.RelativeHorizontalAlignment);
    Console.WriteLine(table.RelativeVerticalAlignment);
}
else
{
    Console.WriteLine(table.Alignment);
}
