Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class CloneTable
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()
        CloneCompleteTable(dataDir)
        CloneLastRow(dataDir)
    End Sub

    Public Shared Sub CloneCompleteTable(dataDir As String)
        ' ExStart:CloneCompleteTable
        Dim doc As New Document(dataDir & Convert.ToString("Table.SimpleTable.doc"))

        ' Retrieve the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        ' Create a clone of the table.
        Dim tableClone As Table = DirectCast(table.Clone(True), Table)

        ' Insert the cloned table into the document after the original
        table.ParentNode.InsertAfter(tableClone, table)

        ' Insert an empty paragraph between the two tables or else they will be combined into one
        ' upon save. This has to do with document validation.
        table.ParentNode.InsertAfter(New Paragraph(doc), table)
        dataDir = dataDir & Convert.ToString("Table.CloneTableAndInsert_out_.doc")

        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:CloneCompleteTable
        Console.WriteLine(Convert.ToString(vbLf & "Table cloned successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub CloneLastRow(dataDir As String)
        ' ExStart:CloneLastRow
        Dim doc As New Document(dataDir & Convert.ToString("Table.SimpleTable.doc"))

        ' Retrieve the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        ' Clone the last row in the table.
        Dim clonedRow As Row = DirectCast(table.LastRow.Clone(True), Row)

        ' Remove all content from the cloned row's cells. This makes the row ready for
        ' new content to be inserted into.
        For Each cell As Cell In clonedRow.Cells
            cell.RemoveAllChildren()
        Next

        ' Add the row to the end of the table.
        table.AppendChild(clonedRow)

        dataDir = dataDir & Convert.ToString("Table.AddCloneRowToTable_out_.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:CloneLastRow
        Console.WriteLine(Convert.ToString(vbLf & "Table last row cloned successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
