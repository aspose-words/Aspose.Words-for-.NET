// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

Cell cell = builder.InsertCell();
builder.Writeln("Outer Table Cell 1");

builder.InsertCell();
builder.Writeln("Outer Table Cell 2");

// This call is important to create a nested table within the first table. 
// Without this call, the cells inserted below will be appended to the outer table.
builder.EndTable();

// Move to the first cell of the outer table.
builder.MoveTo(cell.FirstParagraph);

// Build the inner table.
builder.InsertCell();
builder.Writeln("Inner Table Cell 1");
builder.InsertCell();
builder.Writeln("Inner Table Cell 2");
builder.EndTable();

doc.Save(ArtifactsDir + "WorkingWithTables.NestedTable.docx");
