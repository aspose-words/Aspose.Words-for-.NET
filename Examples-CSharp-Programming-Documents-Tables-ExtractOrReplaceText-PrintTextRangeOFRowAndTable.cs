// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// Print the contents of the second row to the screen.
Console.WriteLine("\nContents of the row: ");
Console.WriteLine(table.Rows[1].Range.Text);

// Print the contents of the last cell in the table to the screen.
Console.WriteLine("\nContents of the cell: ");
Console.WriteLine(table.LastRow.LastCell.Range.Text);
