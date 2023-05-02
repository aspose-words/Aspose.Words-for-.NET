// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
            
// We start by creating the table object. Note that we must pass the document object
// to the constructor of each node. This is because every node we create must belong
// to some document.
Table table = new Table(doc);
doc.FirstSection.Body.AppendChild(table);

// Here we could call EnsureMinimum to create the rows and cells for us. This method is used
// to ensure that the specified node is valid. In this case, a valid table should have at least one Row and one cell.

// Instead, we will handle creating the row and table ourselves.
// This would be the best way to do this if we were creating a table inside an algorithm.
Row row = new Row(doc);
row.RowFormat.AllowBreakAcrossPages = true;
table.AppendChild(row);

// We can now apply any auto fit settings.
table.AutoFit(AutoFitBehavior.FixedColumnWidths);

Cell cell = new Cell(doc);
cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
cell.CellFormat.Width = 80;
cell.AppendChild(new Paragraph(doc));
cell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 1 Text"));

row.AppendChild(cell);

// We would then repeat the process for the other cells and rows in the table.
// We can also speed things up by cloning existing cells and rows.
row.AppendChild(cell.Clone(false));
row.LastCell.AppendChild(new Paragraph(doc));
row.LastCell.FirstParagraph.AppendChild(new Run(doc, "Row 1, Cell 2 Text"));
            
doc.Save(ArtifactsDir + "WorkingWithTables.InsertTableDirectly.docx");
