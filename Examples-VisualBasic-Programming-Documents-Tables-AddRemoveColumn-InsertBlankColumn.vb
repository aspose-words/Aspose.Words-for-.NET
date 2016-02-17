' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Get the first table in the document.
Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

' Get the second column in the table.
Dim column__1 As Column = Column.FromIndex(table, 0)
' Print the plain text of the column to the screen.
Console.WriteLine(column__1.ToTxt())
' Create a new column to the left of this column.
' This is the same as using the "Insert Column Before" command in Microsoft Word.
Dim newColumn As Column = column__1.InsertColumnBefore()

' Add some text to each of the column cells.
For Each cell As Cell In newColumn.Cells
    cell.FirstParagraph.AppendChild(New Run(doc, "Column Text " + newColumn.IndexOf(cell)))
Next
