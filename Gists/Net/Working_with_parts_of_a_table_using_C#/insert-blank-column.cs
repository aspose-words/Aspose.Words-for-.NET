// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

Column column = Column.FromIndex(table, 0);
// Print the plain text of the column to the screen.
Console.WriteLine(column.ToTxt());

            
// Create a new column to the left of this column.
// This is the same as using the "Insert Column Before" command in Microsoft Word.
Column newColumn = column.InsertColumnBefore();

foreach (Cell cell in newColumn.Cells)
    cell.FirstParagraph.AppendChild(new Run(doc, "Column Text " + newColumn.IndexOf(cell)));
