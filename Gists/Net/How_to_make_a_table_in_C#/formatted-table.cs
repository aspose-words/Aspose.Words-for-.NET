// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

Table table = builder.StartTable();
builder.InsertCell();

// Table wide formatting must be applied after at least one row is present in the table.
table.LeftIndent = 20.0;

// Set height and define the height rule for the header row.
builder.RowFormat.Height = 40.0;
builder.RowFormat.HeightRule = HeightRule.AtLeast;

builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
builder.Font.Size = 16;
builder.Font.Name = "Arial";
builder.Font.Bold = true;

builder.CellFormat.Width = 100.0;
builder.Write("Header Row,\n Cell 1");

// We don't need to specify this cell's width because it's inherited from the previous cell.
builder.InsertCell();
builder.Write("Header Row,\n Cell 2");

builder.InsertCell();
builder.CellFormat.Width = 200.0;
builder.Write("Header Row,\n Cell 3");
builder.EndRow();

builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
builder.CellFormat.Width = 100.0;
builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

// Reset height and define a different height rule for table body.
builder.RowFormat.Height = 30.0;
builder.RowFormat.HeightRule = HeightRule.Auto;
builder.InsertCell();
            
// Reset font formatting.
builder.Font.Size = 12;
builder.Font.Bold = false;

builder.Write("Row 1, Cell 1 Content");
builder.InsertCell();
builder.Write("Row 1, Cell 2 Content");

builder.InsertCell();
builder.CellFormat.Width = 200.0;
builder.Write("Row 1, Cell 3 Content");
builder.EndRow();

builder.InsertCell();
builder.CellFormat.Width = 100.0;
builder.Write("Row 2, Cell 1 Content");

builder.InsertCell();
builder.Write("Row 2, Cell 2 Content");

builder.InsertCell();
builder.CellFormat.Width = 200.0;
builder.Write("Row 2, Cell 3 Content.");
builder.EndRow();
builder.EndTable();

doc.Save(ArtifactsDir + "WorkingWithTables.FormattedTable.docx");
