// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Table with merged cells.docx");

Table table = doc.FirstSection.Body.Tables[0];
// Now merged cells have appropriate merge flags.
table.ConvertToHorizontallyMergedCells();
