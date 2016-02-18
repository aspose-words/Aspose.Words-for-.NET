Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class JoiningAndSplittingTable
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()
        Dim fileName As String = "Table.Document.doc"
        CombineRows(dataDir, fileName)
        SplitTable(dataDir, fileName)
    End Sub
    ''' <summary>
    ''' Shows how to combine the rows from two tables into one.
    ''' </summary>        
    Private Shared Sub CombineRows(dataDir As String, fileName As String)
        ' ExStart:CombineRows
        ' Load the document.
        Dim doc As New Document(dataDir & fileName)

        ' Get the first and second table in the document.
        ' The rows from the second table will be appended to the end of the first table.
        Dim firstTable As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
        Dim secondTable As Table = DirectCast(doc.GetChild(NodeType.Table, 1, True), Table)

        ' Append all rows from the current table to the next.
        ' Due to the design of tables even tables with different cell count and widths can be joined into one table.
        While secondTable.HasChildNodes
            firstTable.Rows.Add(secondTable.FirstRow)
        End While

        ' Remove the empty table container.
        secondTable.Remove()
        dataDir = dataDir & Convert.ToString("Table.CombineTables_out_.doc")
        ' Save the finished document.
        doc.Save(dataDir)
        ' ExEnd:CombineRows
        Console.WriteLine(Convert.ToString(vbLf & "Rows combine successfully from two tables into one." & vbLf & "File saved at ") & dataDir)

    End Sub
    ''' <summary>
    ''' Shows how to split a table into two tables in a specific row.
    ''' </summary>              
    Private Shared Sub SplitTable(dataDir As String, fileName As String)
        ' ExStart:SplitTable
        ' Load the document.
        Dim doc As New Document(dataDir & fileName)

        ' Get the first table in the document.
        Dim firstTable As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        ' We will split the table at the third row (inclusive).
        Dim row As Row = firstTable.Rows(2)

        ' Create a new container for the split table.
        Dim table As Table = DirectCast(firstTable.Clone(False), Table)

        ' Insert the container after the original.
        firstTable.ParentNode.InsertAfter(table, firstTable)

        ' Add a buffer paragraph to ensure the tables stay apart.
        firstTable.ParentNode.InsertAfter(New Paragraph(doc), firstTable)

        Dim currentRow As Row

        Do
            currentRow = firstTable.LastRow
            table.PrependChild(currentRow)
        Loop While currentRow IsNot row

        dataDir = dataDir & Convert.ToString("Table.SplitTable_out_.doc")
        ' Save the finished document.
        doc.Save(dataDir)
        ' ExEnd:SplitTable
        Console.WriteLine(Convert.ToString(vbLf & "Table splitted successfully into two tables." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
