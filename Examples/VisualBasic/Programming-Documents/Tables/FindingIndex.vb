Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class FindingIndex
    Public Shared Sub Run()

        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables() + "Table.SimpleTable.doc"
        Dim doc As New Document(dataDir)

        'ExStart:RetrieveTableIndex
        ' Get the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        Dim allTables As NodeCollection = doc.GetChildNodes(NodeType.Table, True)
        Dim tableIndex As Integer = allTables.IndexOf(table)
        'ExEnd:RetrieveTableIndex
        Console.WriteLine(vbLf & "Table index is " + tableIndex.ToString())

        'ExStart:RetrieveRowIndex
        Dim rowIndex As Integer = table.IndexOf(DirectCast(table.LastRow, Row))
        'ExEnd:RetrieveRowIndex
        Console.WriteLine(vbLf & "Row index is " + rowIndex.ToString())

        Dim row As Row = DirectCast(table.LastRow, Row)
        'ExStart:RetrieveCellIndex
        Dim cellIndex As Integer = row.IndexOf(row.Cells(4))
        'ExEnd:RetrieveCellIndex
        Console.WriteLine(vbLf & "Cell index is " + cellIndex.ToString())

    End Sub
End Class
