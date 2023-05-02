// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");
            
Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

// The range text will include control characters such as "\a" for a cell.
// You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.

Console.WriteLine("Contents of the table: ");
Console.WriteLine(table.Range.Text);
