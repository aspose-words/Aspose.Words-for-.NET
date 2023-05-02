// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

builder.InsertCell();
builder.CellFormat.VerticalMerge = CellMerge.First;
builder.Write("Text in merged cells.");

builder.InsertCell();
builder.CellFormat.VerticalMerge = CellMerge.None;
builder.Write("Text in one cell");
builder.EndRow();

builder.InsertCell();
// This cell is vertically merged to the cell above and should be empty.
builder.CellFormat.VerticalMerge = CellMerge.Previous;

builder.InsertCell();
builder.CellFormat.VerticalMerge = CellMerge.None;
builder.Write("Text in another cell");
builder.EndRow();
builder.EndTable();
            
doc.Save(ArtifactsDir + "WorkingWithTables.VerticalMerge.docx");
