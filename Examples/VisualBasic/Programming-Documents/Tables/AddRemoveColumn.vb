Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports System.Text
Imports System.Collections
Public Class AddRemoveColumn
    Public Shared Sub Run()

        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables() + "Table.Document.doc"
        Dim doc As New Document(dataDir)
        InsertBlankColumn(doc)
        RemoveColumn(doc)

    End Sub
    Private Shared Sub RemoveColumn(doc As Document)
        'ExStart:RemoveColumn
        ' Get the second table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 1, True), Table)

        ' Get the third column from the table and remove it.
        Dim column__1 As Column = Column.FromIndex(table, 2)
        column__1.Remove()
        'ExEnd:RemoveColumn
        Console.WriteLine(vbLf & "Third column removed successfully.")
    End Sub
    Private Shared Sub InsertBlankColumn(doc As Document)
        'ExStart:InsertBlankColumn
        ' Get the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        'ExStart:GetPlainText
        ' Get the second column in the table.
        Dim column__1 As Column = Column.FromIndex(table, 0)
        ' Print the plain text of the column to the screen.
        Console.WriteLine(column__1.ToTxt())
        'ExEnd:GetPlainText
        ' Create a new column to the left of this column.
        ' This is the same as using the "Insert Column Before" command in Microsoft Word.
        Dim newColumn As Column = column__1.InsertColumnBefore()

        ' Add some text to each of the column cells.
        For Each cell As Cell In newColumn.Cells
            cell.FirstParagraph.AppendChild(New Run(doc, "Column Text " + newColumn.IndexOf(cell).ToString()))
        Next
        'ExEnd:InsertBlankColumn
        Console.WriteLine(vbLf & "Column added successfully.")
    End Sub
    'ExStart:ColumnClass
    ''' <summary>
    ''' Represents a facade object for a column of a table in a Microsoft Word document.
    ''' </summary>
    Friend Class Column
        Private Sub New(table As Table, columnIndex As Integer)
            If table Is Nothing Then
                Throw New ArgumentException("table")
            End If

            mTable = table
            mColumnIndex = columnIndex
        End Sub

        ''' <summary>
        ''' Returns a new column facade from the table and supplied zero-based index.
        ''' </summary>
        Public Shared Function FromIndex(table As Table, columnIndex As Integer) As Column
            Return New Column(table, columnIndex)
        End Function

        ''' <summary>
        ''' Returns the cells which make up the column.
        ''' </summary>
        Public ReadOnly Property Cells() As Cell()
            Get
                Return DirectCast(GetColumnCells().ToArray(GetType(Cell)), Cell())
            End Get
        End Property

        ''' <summary>
        ''' Returns the index of the given cell in the column.
        ''' </summary>
        Public Function IndexOf(cell As Cell) As Integer
            Return GetColumnCells().IndexOf(cell)
        End Function

        ''' <summary>
        ''' Inserts a brand new column before this column into the table.
        ''' </summary>
        Public Function InsertColumnBefore() As Column
            Dim columnCells As Cell() = Cells

            If columnCells.Length = 0 Then
                Throw New ArgumentException("Column must not be empty")
            End If

            ' Create a clone of this column.
            For Each cell As Cell In columnCells
                cell.ParentRow.InsertBefore(cell.Clone(False), cell)
            Next

            ' This is the new column.
            Dim column As New Column(columnCells(0).ParentRow.ParentTable, mColumnIndex)

            ' We want to make sure that the cells are all valid to work with (have at least one paragraph).
            For Each cell As Cell In column.Cells
                cell.EnsureMinimum()
            Next

            ' Increase the index which this column represents since there is now one extra column infront.
            mColumnIndex += 1

            Return column
        End Function

        ''' <summary>
        ''' Removes the column from the table.
        ''' </summary>
        Public Sub Remove()
            For Each cell As Cell In Cells
                cell.Remove()
            Next
        End Sub

        ''' <summary>
        ''' Returns the text of the column. 
        ''' </summary>
        Public Function ToTxt() As String
            Dim builder As New StringBuilder()

            For Each cell As Cell In Cells
                builder.Append(cell.ToString(SaveFormat.Text))
            Next

            Return builder.ToString()
        End Function

        ''' <summary>
        ''' Provides an up-to-date collection of cells which make up the column represented by this facade.
        ''' </summary>
        Private Function GetColumnCells() As ArrayList
            Dim columnCells As New ArrayList()

            For Each row As Row In mTable.Rows
                Dim cell As Cell = row.Cells(mColumnIndex)
                If cell IsNot Nothing Then
                    columnCells.Add(cell)
                End If
            Next

            Return columnCells
        End Function

        Private mColumnIndex As Integer
        Private mTable As Table
    End Class
    'ExEnd:ColumnClass
End Class
