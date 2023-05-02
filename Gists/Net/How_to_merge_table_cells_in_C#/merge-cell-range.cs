// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Table with merged cells.docx");

Table table = doc.FirstSection.Body.Tables[0];

// We want to merge the range of cells found inbetween these two cells.
Cell cellStartRange = table.Rows[0].Cells[0];
Cell cellEndRange = table.Rows[1].Cells[1];

// Merge all the cells between the two specified cells into one.
MergeCells(cellStartRange, cellEndRange);
            
doc.Save(ArtifactsDir + "WorkingWithTables.MergeCellRange.docx");
