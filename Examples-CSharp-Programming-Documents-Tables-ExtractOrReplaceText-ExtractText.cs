// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir);

// Get the first table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// The range text will include control characters such as "\a" for a cell.
// You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.

// Print the plain text range of the table to the screen.
Console.WriteLine("Contents of the table: ");
Console.WriteLine(table.Range.Text);
