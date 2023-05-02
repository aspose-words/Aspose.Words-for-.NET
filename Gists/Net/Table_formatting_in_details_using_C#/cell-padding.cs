// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

builder.StartTable();
builder.InsertCell();

// Sets the amount of space (in points) to add to the left/top/right/bottom of the cell's contents.
builder.CellFormat.SetPaddings(30, 50, 30, 50);
builder.Writeln("I'm a wonderful formatted cell.");

builder.EndRow();
builder.EndTable();

doc.Save(ArtifactsDir + "WorkingWithTableStylesAndFormatting.CellPadding.docx");
