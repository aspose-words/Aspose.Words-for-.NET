// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Insert a table with a width that takes up half the page width.
Table table = builder.StartTable();

builder.InsertCell();
table.PreferredWidth = PreferredWidth.FromPercent(50);
builder.Writeln("Cell #1");

builder.InsertCell();
builder.Writeln("Cell #2");

builder.InsertCell();
builder.Writeln("Cell #3");

doc.Save(ArtifactsDir + "WorkingWithTables.AutoFitPageWidth.docx");
