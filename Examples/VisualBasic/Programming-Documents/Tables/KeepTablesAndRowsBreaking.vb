
Imports System.IO
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports System.Diagnostics
Public Class KeepTablesAndRowsBreaking
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()

        ' The below method shows how to disable rows breaking across pages for every row in a table.
        RowFormatDisableBreakAcrossPages(dataDir)
        ' The below method shows how to set a table to stay together on the same page.
        KeepTableTogether(dataDir)

    End Sub
    Public Shared Sub RowFormatDisableBreakAcrossPages(dataDir As String)
        ' ExStart:RowFormatDisableBreakAcrossPages
        Dim doc As New Document(dataDir & Convert.ToString("Table.TableAcrossPage.doc"))

        ' Retrieve the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
        ' Disable breaking across pages for all rows in the table.
        For Each row As Row In table
            row.RowFormat.AllowBreakAcrossPages = False
        Next

        dataDir = dataDir & Convert.ToString("Table.DisableBreakAcrossPages_out.doc")
        doc.Save(dataDir)
        ' ExEnd:RowFormatDisableBreakAcrossPages
        Console.WriteLine(Convert.ToString(vbLf & "Table rows breaking across pages for every row in a table disabled successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub KeepTableTogether(dataDir As String)
        ' ExStart:KeepTableTogether
        Dim doc As New Document(dataDir & Convert.ToString("Table.TableAcrossPage.doc"))
        ' Retrieve the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        ' To keep a table from breaking across a page we need to enable KeepWithNext 
        ' For every paragraph in the table except for the last paragraphs in the last 
        ' Row of the table.
        For Each cell As Cell In table.GetChildNodes(NodeType.Cell, True)
            ' Call this method if table' S cell is created on the fly
            ' Newly created cell does not have paragraph inside
            cell.EnsureMinimum()
            For Each para As Paragraph In cell.Paragraphs
                If Not (cell.ParentRow.IsLastRow AndAlso para.IsEndOfCell) Then
                    para.ParagraphFormat.KeepWithNext = True
                End If
            Next
        Next
        dataDir = dataDir & Convert.ToString("Table.KeepTableTogether_out.doc")
        doc.Save(dataDir)
        ' ExEnd:KeepTableTogether
        Console.WriteLine(Convert.ToString(vbLf & "Table setup successfully to stay together on the same page." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
