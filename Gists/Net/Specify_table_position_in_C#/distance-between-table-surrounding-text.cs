// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");

Console.WriteLine("\nGet distance between table left, right, bottom, top and the surrounding text.");
Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

Console.WriteLine(table.DistanceTop);
Console.WriteLine(table.DistanceBottom);
Console.WriteLine(table.DistanceRight);
Console.WriteLine(table.DistanceLeft);
