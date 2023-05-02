// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
            
// Start building the table.
builder.StartTable();
builder.InsertCell();
builder.Write("Row 1, Cell 1 Content.");
            
// Build the second cell.
builder.InsertCell();
builder.Write("Row 1, Cell 2 Content.");
            
// Call the following method to end the row and start a new row.
builder.EndRow();

// Build the first cell of the second row.
builder.InsertCell();
builder.Write("Row 2, Cell 1 Content");

// Build the second cell.
builder.InsertCell();
builder.Write("Row 2, Cell 2 Content.");
builder.EndRow();

// Signal that we have finished building the table.
builder.EndTable();

doc.Save(ArtifactsDir + "WorkingWithTables.CreateSimpleTable.docx");
