' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Text

Imports Aspose.Words
Imports Aspose.Words.Tables

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExTableColumn
		Inherits ApiExampleBase
		'ExStart
		'ExId:ColumnFacade
		'ExSummary:Demonstrates a facade object for working with a column of a table.
		''' <summary>
		''' Represents a facade object for a column of a table in a Microsoft Word document.
		''' </summary>
		Public Class Column
			Private Sub New(ByVal table As Table, ByVal columnIndex As Integer)
				If table Is Nothing Then
					Throw New ArgumentException("table")
				End If

				Me.mTable = table
				Me.mColumnIndex = columnIndex
			End Sub

			''' <summary>
			''' Returns a new column facade from the table and supplied zero-based index.
			''' </summary>
			Public Shared Function FromIndex(ByVal table As Table, ByVal columnIndex As Integer) As Column
				Return New Column(table, columnIndex)
			End Function

			''' <summary>
			''' Returns the cells which make up the column.
			''' </summary>
			Public ReadOnly Property Cells() As Cell()
				Get
					Return CType(Me.GetColumnCells().ToArray(GetType(Cell)), Cell())
				End Get
			End Property

			''' <summary>
			''' Returns the index of the given cell in the column.
			''' </summary>
			Public Function IndexOf(ByVal cell As Cell) As Integer
				Return Me.GetColumnCells().IndexOf(cell)
			End Function

			''' <summary>
			''' Inserts a brand new column before this column into the table.
			''' </summary>
			Public Function InsertColumnBefore() As Column
				Dim columnCells() As Cell = Me.Cells

				If columnCells.Length = 0 Then
					Throw New ArgumentException("Column must not be empty")
				End If

				' Create a clone of this column.
				For Each cell As Cell In columnCells
					cell.ParentRow.InsertBefore(cell.Clone(False), cell)
				Next cell

				' This is the new column.
				Dim column As New Column(columnCells(0).ParentRow.ParentTable, Me.mColumnIndex)

				' We want to make sure that the cells are all valid to work with (have at least one paragraph).
				For Each cell As Cell In column.Cells
					cell.EnsureMinimum()
				Next cell

				' Increase the index which this column represents since there is now one extra column infront.
				Me.mColumnIndex += 1

				Return column
			End Function

			''' <summary>
			''' Removes the column from the table.
			''' </summary>
			Public Sub Remove()
				For Each cell As Cell In Me.Cells
					cell.Remove()
				Next cell
			End Sub

			''' <summary>
			''' Returns the text of the column. 
			''' </summary>
			Public Function ToTxt() As String
				Dim builder As New StringBuilder()

				For Each cell As Cell In Me.Cells
					builder.Append(cell.ToString(SaveFormat.Text))
				Next cell

				Return builder.ToString()
			End Function

			''' <summary>
			''' Provides an up-to-date collection of cells which make up the column represented by this facade.
			''' </summary>
			Private Function GetColumnCells() As ArrayList
				Dim columnCells As New ArrayList()

				For Each row As Row In Me.mTable.Rows
					Dim cell As Cell = row.Cells(Me.mColumnIndex)
					If cell IsNot Nothing Then
						columnCells.Add(cell)
					End If
				Next row

				Return columnCells
			End Function

			Private mColumnIndex As Integer
			Private mTable As Table
		End Class
		'ExEnd

		<Test> _
		Public Sub RemoveColumnFromTable()
			'ExStart
			'ExId:RemoveTableColumn
			'ExSummary:Shows how to remove a column from a table in a document.
			Dim doc As New Document(MyDir & "Table.Document.doc")
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 1, True), Table)

			' Get the third column from the table and remove it.
			Dim column As Column = Column.FromIndex(table, 2)
			column.Remove()
			'ExEnd

			doc.Save(MyDir & "\Artifacts\Table.RemoveColumn.doc")

			Assert.AreEqual(16, table.GetChildNodes(NodeType.Cell, True).Count)
			Assert.AreEqual("Cell 3 contents", table.Rows(2).Cells(2).ToString(SaveFormat.Text).Trim())
			Assert.AreEqual("Cell 3 contents", table.LastRow.Cells(2).ToString(SaveFormat.Text).Trim())
		End Sub

		<Test> _
		Public Sub InsertNewColumnIntoTable()
			Dim doc As New Document(MyDir & "Table.Document.doc")
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 1, True), Table)

			'ExStart
			'ExId:InsertNewColumn
			'ExSummary:Shows how to insert a blank column into a table.
			' Get the second column in the table.
			Dim column As Column = Column.FromIndex(table, 1)

			' Create a new column to the left of this column.
			' This is the same as using the "Insert Column Before" command in Microsoft Word.
			Dim newColumn As Column = column.InsertColumnBefore()

			' Add some text to each of the column cells.
			For Each cell As Cell In newColumn.Cells
				cell.FirstParagraph.AppendChild(New Run(doc, "Column Text " & newColumn.IndexOf(cell)))
			Next cell
			'ExEnd

			doc.Save(MyDir & "\Artifacts\Table.InsertColumn.doc")

			Assert.AreEqual(24, table.GetChildNodes(NodeType.Cell, True).Count)
			Assert.AreEqual("Column Text 0", table.FirstRow.Cells(1).ToString(SaveFormat.Text).Trim())
			Assert.AreEqual("Column Text 3", table.LastRow.Cells(1).ToString(SaveFormat.Text).Trim())
		End Sub

		<Test> _
		Public Sub TableColumnToTxt()
			Dim doc As New Document(MyDir & "Table.Document.doc")
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 1, True), Table)

			'ExStart
			'ExId:TableColumnToTxt
			'ExSummary:Shows how to get the plain text of a table column.
			' Get the first column in the table.
			Dim column As Column = Column.FromIndex(table, 0)

			' Print the plain text of the column to the screen.
			Console.WriteLine(column.ToTxt())
			'ExEnd

			Assert.AreEqual(Constants.vbCrLf & "Row 1" & Constants.vbCrLf & "Row 2" & Constants.vbCrLf & "Row 3" & Constants.vbCrLf, column.ToTxt())
		End Sub
	End Class
End Namespace
