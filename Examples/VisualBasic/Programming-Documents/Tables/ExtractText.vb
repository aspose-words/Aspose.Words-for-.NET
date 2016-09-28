Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports Aspose.Words.Replacing
Public Class ExtractText
    Public Shared Sub Run()

        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables() + "Table.SimpleTable.doc"
        ExtractPrintText(dataDir)
        ReplaceText(dataDir)

    End Sub
    Private Shared Sub ExtractPrintText(dataDir As String)
        'ExStart:ExtractText
        Dim doc As New Document(dataDir)

        ' Get the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        ' The range text will include control characters such as "\a" for a cell.
        ' You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.

        ' Print the plain text range of the table to the screen.
        Console.WriteLine("Contents of the table: ")
        Console.WriteLine(table.Range.Text)
        'ExEnd:ExtractText   

        'ExStart:PrintTextRangeOFRowAndTable
        ' Print the contents of the second row to the screen.
        Console.WriteLine(vbLf & "Contents of the row: ")
        Console.WriteLine(table.Rows(1).Range.Text)

        ' Print the contents of the last cell in the table to the screen.
        Console.WriteLine(vbLf & "Contents of the cell: ")
        Console.WriteLine(table.LastRow.LastCell.Range.Text)
        'ExEnd:PrintTextRangeOFRowAndTable
    End Sub
    Private Shared Sub ReplaceText(dataDir As String)
        'ExStart:ReplaceText
        Dim doc As New Document(dataDir)

        ' Get the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        ' Replace any instances of our string in the entire table.
        table.Range.Replace("Carrots", "Eggs", New FindReplaceOptions(FindReplaceDirection.Forward))
        ' Replace any instances of our string in the last cell of the table only.
        table.LastRow.LastCell.Range.Replace("50", "20", New FindReplaceOptions(FindReplaceDirection.Forward))

        dataDir = RunExamples.GetDataDir_WorkingWithTables() + "Table.ReplaceCellText_out_.doc"
        doc.Save(dataDir)
        'ExEnd:ReplaceText    
        Console.WriteLine(Convert.ToString(vbLf & "Text replaced successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
