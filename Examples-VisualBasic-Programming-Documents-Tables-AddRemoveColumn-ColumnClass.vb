' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
