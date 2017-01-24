' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing

Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Tables

Imports NUnit.Framework

Namespace ApiExamples
	''' <summary>
	''' Examples using tables in documents.
	''' </summary>
	<TestFixture> _
	Public Class ExTable
		Inherits ApiExampleBase
		<Test> _
		Public Sub DisplayContentOfTables()
			'ExStart
			'ExFor:Table
			'ExFor:Row.Cells
			'ExFor:Table.Rows
			'ExFor:Cell
			'ExFor:Row
			'ExFor:RowCollection
			'ExFor:CellCollection
			'ExFor:NodeCollection.IndexOf(Node)
			'ExSummary:Shows how to iterate through all tables in the document and display the content from each cell.
			Dim doc As New Document(MyDir & "Table.Document.doc")

			' Here we get all tables from the Document node. You can do this for any other composite node
			' which can contain block level nodes. For example you can retrieve tables from header or from a cell
			' containing another table (nested tables).
			Dim tables As NodeCollection = doc.GetChildNodes(NodeType.Table, True)

			' Iterate through all tables in the document
			For Each table As Table In tables
				' Get the index of the table node as contained in the parent node of the table
				Dim tableIndex As Integer = table.ParentNode.ChildNodes.IndexOf(table)
				Console.WriteLine("Start of Table {0}", tableIndex)

				' Iterate through all rows in the table
				For Each row As Row In table.Rows
					Dim rowIndex As Integer = table.Rows.IndexOf(row)
					Console.WriteLine(Constants.vbTab & "Start of Row {0}", rowIndex)

					' Iterate through all cells in the row
					For Each cell As Cell In row.Cells
						Dim cellIndex As Integer = row.Cells.IndexOf(cell)
						' Get the plain text content of this cell.
						Dim cellText As String = cell.ToString(SaveFormat.Text).Trim()
						' Print the content of the cell.
						Console.WriteLine(Constants.vbTab + Constants.vbTab & "Contents of Cell:{0} = ""{1}""", cellIndex, cellText)
					Next cell
					'Console.WriteLine();
					Console.WriteLine(Constants.vbTab & "End of Row {0}", rowIndex)
				Next row
				Console.WriteLine("End of Table {0}", tableIndex)
				Console.WriteLine()
			Next table
			'ExEnd

			Assert.Greater(tables.Count, 0)
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub CalcuateDepthOfNestedTablesCaller()
			Me.CalcuateDepthOfNestedTables()
		End Sub

		'ExStart
		'ExFor:Node.GetAncestor(NodeType)
		'ExFor:Table.NodeType
		'ExFor:Cell.Tables
		'ExFor:TableCollection
		'ExFor:NodeCollection.Count
		'ExSummary:Shows how to find out if a table contains another table or if the table itself is nested inside another table.
		Public Sub CalcuateDepthOfNestedTables()
			Dim doc As New Document(MyDir & "Table.NestedTables.doc")
			Dim tableIndex As Integer = 0

			For Each table As Table In doc.GetChildNodes(NodeType.Table, True)
				' First lets find if any cells in the table have tables themselves as children.
				Dim count As Integer = GetChildTableCount(table)
				Console.WriteLine("Table #{0} has {1} tables directly within its cells", tableIndex, count)

				' Now let's try the other way around, lets try find if the table is nested inside another table and at what depth.
				Dim tableDepth As Integer = GetNestedDepthOfTable(table)

				If tableDepth > 0 Then
					Console.WriteLine("Table #{0} is nested inside another table at depth of {1}", tableIndex, tableDepth)
				Else
					Console.WriteLine("Table #{0} is a non nested table (is not a child of another table)", tableIndex)
				End If

				tableIndex += 1
			Next table
		End Sub

		''' <summary>
		''' Calculates what level a table is nested inside other tables.
		''' <returns>
		''' An integer containing the level the table is nested at.
		''' 0 = Table is not nested inside any other table
		''' 1 = Table is nested within one parent table
		''' 2 = Table is nested within two parent tables etc..</returns>
		''' </summary>
		Private Shared Function GetNestedDepthOfTable(ByVal table As Table) As Integer
			Dim depth As Integer = 0

			Dim type As NodeType = table.NodeType
			' The parent of the table will be a Cell, instead attempt to find a grandparent that is of type Table
			Dim parent As Node = table.GetAncestor(type)

			Do While parent IsNot Nothing
				' Every time we find a table a level up we increase the depth counter and then try to find an
				' ancestor of type table from the parent.
				depth += 1
				parent = parent.GetAncestor(type)
			Loop

			Return depth
		End Function

		''' <summary>
		''' Determines if a table contains any immediate child table within its cells.
		''' Does not recursively traverse through those tables to check for further tables.
		''' <returns>Returns true if at least one child cell contains a table.
		''' Returns false if no cells in the table contains a table.</returns>
		''' </summary>
		Private Shared Function GetChildTableCount(ByVal table As Table) As Integer
			Dim tableCount As Integer = 0
			' Iterate through all child rows in the table
			For Each row As Row In table.Rows
				' Iterate through all child cells in the row
				For Each Cell As Cell In row.Cells
					' Retrieve the collection of child tables of this cell
					Dim childTables As TableCollection = Cell.Tables

					' If this cell has a table as a child then return true
					If childTables.Count > 0 Then
						tableCount += 1
					End If
				Next Cell
			Next row

			' No cell contains a table
			Return tableCount
		End Function
		'ExEnd

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub ConvertTextboxToTableCaller()
			Me.ConvertTextboxToTable()
		End Sub

		'ExStart
		'ExId:TextboxToTable
		'ExSummary:Shows how to convert a textbox to a table and retain almost identical formatting. This is useful for HTML export.
		Public Sub ConvertTextboxToTable()
			' Open the document
			Dim doc As New Document(MyDir & "Shape.Textbox.doc")

			' Convert all shape nodes which contain child nodes.
			' We convert the collection to an array as static "snapshot" because the original textboxes will be removed after conversion which will
			' invalidate the enumerator.
			For Each shape As Shape In doc.GetChildNodes(NodeType.Shape, True).ToArray()
				If shape.HasChildNodes Then
					ConvertTextboxToTable(shape)
				End If
			Next shape

			doc.Save(MyDir & "\Artifacts\Table.ConvertTextboxToTable.html")
		End Sub

		''' <summary>
		''' Converts a textbox to a table by copying the same content and formatting.
		''' Currently export to HTML will render the textbox as an image which looses any text functionality.
		''' This is useful to convert textboxes in order to retain proper text.
		''' </summary>
		''' <param name="textbox">The textbox shape to convert to a table</param>
		Private Shared Sub ConvertTextboxToTable(ByVal textBox As Shape)
			If textBox.StoryType <> StoryType.Textbox Then
				Throw New ArgumentException("Can only convert a shape of type textbox")
			End If

			Dim doc As Document = CType(textBox.Document, Document)
			Dim section As Section = CType(textBox.GetAncestor(NodeType.Section), Section)

			' Create a table to replace the textbox and transfer the same content and formatting.
			Dim table As New Table(doc)
			' Ensure that the table contains a row and a cell.
			table.EnsureMinimum()
			' Use fixed column widths.
			table.AutoFit(AutoFitBehavior.FixedColumnWidths)

			' A shape is inline level (within a paragraph) where a table can only be block level so insert the table
			' after the paragraph which contains the shape.
			Dim shapeParent As Node = textBox.ParentNode
			shapeParent.ParentNode.InsertAfter(table, shapeParent)

			' If the textbox is not inline then try to match the shape's left position using the table's left indent.
			If (Not textBox.IsInline) AndAlso textBox.Left < section.PageSetup.PageWidth Then
				table.LeftIndent = textBox.Left
			End If

			' We are only using one cell to replicate a textbox so we can make use of the FirstRow and FirstCell property.
			' Carry over borders and shading.
			Dim firstRow As Row = table.FirstRow
			Dim firstCell As Cell = firstRow.FirstCell
			firstCell.CellFormat.Borders.Color = textBox.StrokeColor
			firstCell.CellFormat.Borders.LineWidth = textBox.StrokeWeight
			firstCell.CellFormat.Shading.BackgroundPatternColor = textBox.Fill.Color

			' Transfer the same height and width of the textbox to the table.
			firstRow.RowFormat.HeightRule = HeightRule.Exactly
			firstRow.RowFormat.Height = textBox.Height
			firstCell.CellFormat.Width = textBox.Width
			table.AllowAutoFit = False

			' Replicate the textbox's horizontal alignment.
			Dim horizontalAlignment As TableAlignment
			Select Case textBox.HorizontalAlignment
				Case HorizontalAlignment.Left
					horizontalAlignment = TableAlignment.Left
				Case HorizontalAlignment.Center
					horizontalAlignment = TableAlignment.Center
				Case HorizontalAlignment.Right
					horizontalAlignment = TableAlignment.Right
				Case Else
					' Most other options are left by default.
					horizontalAlignment = TableAlignment.Left

			End Select

			table.Alignment = horizontalAlignment
			firstCell.RemoveAllChildren()

			' Append all content from the textbox to the new table
			For Each node As Node In textBox.GetChildNodes(NodeType.Any, False).ToArray()
				table.FirstRow.FirstCell.AppendChild(node)
			Next node

			' Remove the empty textbox from the document.
			textBox.Remove()
		End Sub
		'ExEnd

		<Test> _
		Public Sub EnsureTableMinimum()
			'ExStart
			'ExFor:Table.EnsureMinimum
			'ExSummary:Shows how to ensure a table node is valid.
			Dim doc As New Document()

			' Create a new table and add it to the document.
			Dim table As New Table(doc)
			doc.FirstSection.Body.AppendChild(table)

			' Ensure the table is valid (has at least one row with one cell).
			table.EnsureMinimum()
			'ExEnd
		End Sub

		<Test> _
		Public Sub EnsureRowMinimum()
			'ExStart
			'ExFor:Row.EnsureMinimum
			'ExSummary:Shows how to ensure a row node is valid.
			Dim doc As New Document()

			' Create a new table and add it to the document.
			Dim table As New Table(doc)
			doc.FirstSection.Body.AppendChild(table)

			' Create a new row and add it to the table.
			Dim row As New Row(doc)
			table.AppendChild(row)

			' Ensure the row is valid (has at least one cell).
			row.EnsureMinimum()
			'ExEnd
		End Sub

		<Test> _
		Public Sub EnsureCellMinimum()
			'ExStart
			'ExFor:Cell.EnsureMinimum
			'ExSummary:Shows how to ensure a cell node is valid.
			Dim doc As New Document(MyDir & "Table.Document.doc")

			' Gets the first cell in the document.
			Dim cell As Cell = CType(doc.GetChild(NodeType.Cell, 0, True), Cell)

			' Ensure the cell is valid (the last child is a paragraph).
			cell.EnsureMinimum()
			'ExEnd
		End Sub

		<Test> _
		Public Sub SetTableBordersOutline()
			'ExStart
			'ExFor:Table.Alignment
			'ExFor:TableAlignment
			'ExFor:Table.ClearBorders
			'ExFor:Table.SetBorder
			'ExFor:TextureIndex
			'ExFor:Table.SetShading
			'ExId:TableBordersOutline
			'ExSummary:Shows how to apply a outline border to a table.
			Dim doc As New Document(MyDir & "Table.EmptyTable.doc")
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' Align the table to the center of the page.
			table.Alignment = TableAlignment.Center

			' Clear any existing borders from the table.
			table.ClearBorders()

			' Set a green border around the table but not inside. 
			table.SetBorder(BorderType.Left, LineStyle.Single, 1.5, Color.Green, True)
			table.SetBorder(BorderType.Right, LineStyle.Single, 1.5, Color.Green, True)
			table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Green, True)
			table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Green, True)

			' Fill the cells with a light green solid color.
			table.SetShading(TextureIndex.TextureSolid, Color.LightGreen, Color.Empty)

			doc.Save(MyDir & "\Artifacts\Table.SetOutlineBorders.doc")
			'ExEnd

			' Verify the borders were set correctly.
			doc = New Document(MyDir & "\Artifacts\Table.SetOutlineBorders.doc")
			Assert.AreEqual(TableAlignment.Center, table.Alignment)
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Top.Color.ToArgb())
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Left.Color.ToArgb())
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Right.Color.ToArgb())
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Bottom.Color.ToArgb())
			Assert.AreNotEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Horizontal.Color.ToArgb())
			Assert.AreNotEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Vertical.Color.ToArgb())
			Assert.AreEqual(Color.LightGreen.ToArgb(), table.FirstRow.FirstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb())
		End Sub

		<Test> _
		Public Sub SetTableBordersAll()
			'ExStart
			'ExFor:Table.SetBorders
			'ExId:TableBordersAll
			'ExSummary:Shows how to build a table with all borders enabled (grid).
			Dim doc As New Document(MyDir & "Table.EmptyTable.doc")
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' Clear any existing borders from the table.
			table.ClearBorders()

			' Set a green border around and inside the table.
			table.SetBorders(LineStyle.Single, 1.5, Color.Green)

			doc.Save(MyDir & "\Artifacts\Table.SetAllBorders.doc")
			'ExEnd

			' Verify the borders were set correctly.
			doc = New Document(MyDir & "\Artifacts\Table.SetAllBorders.doc")
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Top.Color.ToArgb())
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Left.Color.ToArgb())
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Right.Color.ToArgb())
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Bottom.Color.ToArgb())
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Horizontal.Color.ToArgb())
			Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.RowFormat.Borders.Vertical.Color.ToArgb())
		End Sub

		<Test> _
		Public Sub RowFormatProperties()
			'ExStart
			'ExFor:RowFormat
			'ExFor:Row.RowFormat
			'ExId:RowFormatProperties
			'ExSummary:Shows how to modify formatting of a table row.
			Dim doc As New Document(MyDir & "Table.Document.doc")
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' Retrieve the first row in the table.
			Dim firstRow As Row = table.FirstRow

			' Modify some row level properties.
			firstRow.RowFormat.Borders.LineStyle = LineStyle.None
			firstRow.RowFormat.HeightRule = HeightRule.Auto
			firstRow.RowFormat.AllowBreakAcrossPages = True
			'ExEnd

			doc.Save(MyDir & "\Artifacts\Table.RowFormat.doc")

			doc = New Document(MyDir & "\Artifacts\Table.RowFormat.doc")
			table = CType(doc.GetChild(NodeType.Table, 0, True), Table)
			Assert.AreEqual(LineStyle.None, table.FirstRow.RowFormat.Borders.LineStyle)
			Assert.AreEqual(HeightRule.Auto, table.FirstRow.RowFormat.HeightRule)
			Assert.True(table.FirstRow.RowFormat.AllowBreakAcrossPages)
		End Sub

		<Test> _
		Public Sub CellFormatProperties()
			'ExStart
			'ExFor:CellFormat
			'ExFor:Cell.CellFormat
			'ExId:CellFormatProperties
			'ExSummary:Shows how to modify formatting of a table cell.
			Dim doc As New Document(MyDir & "Table.Document.doc")
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' Retrieve the first cell in the table.
			Dim firstCell As Cell = table.FirstRow.FirstCell

			' Modify some row level properties.
			firstCell.CellFormat.Width = 30 ' in points
			firstCell.CellFormat.Orientation = TextOrientation.Downward
			firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen
			'ExEnd

			doc.Save(MyDir & "\Artifacts\Table.CellFormat.doc")

			doc = New Document(MyDir & "\Artifacts\Table.CellFormat.doc")
			table = CType(doc.GetChild(NodeType.Table, 0, True), Table)
			Assert.AreEqual(30, table.FirstRow.FirstCell.CellFormat.Width)
			Assert.AreEqual(TextOrientation.Downward, table.FirstRow.FirstCell.CellFormat.Orientation)
			Assert.AreEqual(Color.LightGreen.ToArgb(), table.FirstRow.FirstCell.CellFormat.Shading.ForegroundPatternColor.ToArgb())
		End Sub

		<Test> _
		Public Sub RemoveBordersFromAllCells()
			'ExStart
			'ExFor:Table
			'ExFor:Table.ClearBorders
			'ExSummary:Shows how to remove all borders from a table.
			Dim doc As New Document(MyDir & "Table.Document.doc")

			' Remove all borders from the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' Clear the borders all cells in the table.
			table.ClearBorders()

			doc.Save(MyDir & "\Artifacts\Table.ClearBorders.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ReplaceTextInTable()
			'ExStart
			'ExFor:Range.Replace(String, String, Boolean, Boolean)
			'ExFor:Cell
			'ExId:ReplaceTextTable
			'ExSummary:Shows how to replace all instances of string of text in a table and cell.
			Dim doc As New Document(MyDir & "Table.SimpleTable.doc")

			' Get the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' Replace any instances of our string in the entire table.
			table.Range.Replace("Carrots", "Eggs", True, True)
			' Replace any instances of our string in the last cell of the table only.
			table.LastRow.LastCell.Range.Replace("50", "20", True, True)

			doc.Save(MyDir & "\Artifacts\Table.ReplaceCellText.doc")
			'ExEnd

			Assert.AreEqual("20", table.LastRow.LastCell.ToString(SaveFormat.Text).Trim())
		End Sub

		<Test> _
		Public Sub PrintTableRange()
			'ExStart
			'ExId:PrintTableRange
			'ExSummary:Shows how to print the text range of a table.
			Dim doc As New Document(MyDir & "Table.SimpleTable.doc")

			' Get the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' The range text will include control characters such as "\a" for a cell.
			' You can call ToString on the desired node to retrieve the plain text content.

			' Print the plain text range of the table to the screen.
			Console.WriteLine("Contents of the table: ")
			Console.WriteLine(table.Range.Text)
			'ExEnd

			'ExStart
			'ExId:PrintRowAndCellRange
			'ExSummary:Shows how to print the text range of row and table elements.
			' Print the contents of the second row to the screen.
			Console.WriteLine(Constants.vbLf & "Contents of the row: ")
			Console.WriteLine(table.Rows(1).Range.Text)

			' Print the contents of the last cell in the table to the screen.
			Console.WriteLine(Constants.vbLf & "Contents of the cell: ")
			Console.WriteLine(table.LastRow.LastCell.Range.Text)
			'ExEnd

			Assert.AreEqual("Apples" & Constants.vbCr + ControlChar.Cell & "20" & Constants.vbCr + ControlChar.Cell + ControlChar.Cell, table.Rows(1).Range.Text)
			Assert.AreEqual("50" & Constants.vbCr & "\a", table.LastRow.LastCell.Range.Text)
		End Sub

		<Test> _
		Public Sub CloneTable()
			'ExStart
			'ExId:CloneTable
			'ExSummary:Shows how to make a clone of a table in the document and insert it after the original table.
			Dim doc As New Document(MyDir & "Table.SimpleTable.doc")

			' Retrieve the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' Create a clone of the table.
			Dim tableClone As Table = CType(table.Clone(True), Table)

			' Insert the cloned table into the document after the original
			table.ParentNode.InsertAfter(tableClone, table)

			' Insert an empty paragraph between the two tables or else they will be combined into one
			' upon save. This has to do with document validation.
			table.ParentNode.InsertAfter(New Paragraph(doc), table)

			doc.Save(MyDir & "\Artifacts\Table.CloneTableAndInsert.doc")
			'ExEnd

			' Verify that the table was cloned and inserted properly.
			Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, True).Count)
			Assert.AreEqual(table.Range.Text, tableClone.Range.Text)

			'ExStart
			'ExId:CloneTableRemoveContent
			'ExSummary:Shows how to remove all content from the cells of a cloned table.
			For Each cell As Cell In tableClone.GetChildNodes(NodeType.Cell, True)
				cell.RemoveAllChildren()
			Next cell
			'ExEnd

			Assert.AreEqual(String.Empty, tableClone.ToString(SaveFormat.Text).Trim())
		End Sub

		<Test> _
		Public Sub RowFormatDisableBreakAcrossPages()
			Dim doc As New Document(MyDir & "Table.TableAcrossPage.doc")

			' Retrieve the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			'ExStart
			'ExFor:RowFormat.AllowBreakAcrossPages
			'ExId:RowFormatAllowBreaks
			'ExSummary:Shows how to disable rows breaking across pages for every row in a table.
			' Disable breaking across pages for all rows in the table.
			For Each row As Row In table
				row.RowFormat.AllowBreakAcrossPages = False
			Next row
			'ExEnd

			doc.Save(MyDir & "\Artifacts\Table.DisableBreakAcrossPages.doc")

			Assert.False(table.FirstRow.RowFormat.AllowBreakAcrossPages)
			Assert.False(table.LastRow.RowFormat.AllowBreakAcrossPages)
		End Sub

		<Test> _
		Public Sub AllowAutoFitOnTable()
			Dim doc As New Document()

			Dim table As New Table(doc)
			table.EnsureMinimum()

			'ExStart
			'ExFor:Table.AllowAutoFit
			'ExId:AllowAutoFit
			'ExSummary:Shows how to set a table to shrink or grow each cell to accommodate its contents.
			table.AllowAutoFit = True
			'ExEnd
		End Sub

		<Test> _
		Public Sub KeepTableTogether()
			Dim doc As New Document(MyDir & "Table.TableAcrossPage.doc")

			' Retrieve the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			'ExStart
			'ExFor:ParagraphFormat.KeepWithNext
			'ExFor:Row.IsLastRow
			'ExFor:Paragraph.IsEndOfCell
			'ExFor:Cell.ParentRow
			'ExFor:Cell.Paragraphs
			'ExId:KeepTableTogether
			'ExSummary:Shows how to set a table to stay together on the same page.
			' To keep a table from breaking across a page we need to enable KeepWithNext 
			' for every paragraph in the table except for the last paragraphs in the last 
			' row of the table.
			For Each cell As Cell In table.GetChildNodes(NodeType.Cell, True)
				For Each para As Paragraph In cell.Paragraphs
					If Not(cell.ParentRow.IsLastRow AndAlso para.IsEndOfCell) Then
						para.ParagraphFormat.KeepWithNext = True
					End If
				Next para
			Next cell
			'ExEnd

			doc.Save(MyDir & "\Artifacts\Table.KeepTableTogether.doc")

			' Verify the correct paragraphs were set properly.
			For Each para As Paragraph In table.GetChildNodes(NodeType.Paragraph, True)
				If para.IsEndOfCell AndAlso (CType(para.ParentNode, Cell)).ParentRow.IsLastRow Then
					Assert.False(para.ParagraphFormat.KeepWithNext)
				Else
					Assert.True(para.ParagraphFormat.KeepWithNext)
				End If
			Next para
		End Sub

		<Test> _
		Public Sub AddClonedRowToTable()
			'ExStart
			'ExFor:Row
			'ExId:AddClonedRowToTable
			'ExSummary:Shows how to make a clone of the last row of a table and append it to the table.
			Dim doc As New Document(MyDir & "Table.SimpleTable.doc")

			' Retrieve the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' Clone the last row in the table.
			Dim clonedRow As Row = CType(table.LastRow.Clone(True), Row)

			' Remove all content from the cloned row's cells. This makes the row ready for
			' new content to be inserted into.
			For Each cell As Cell In clonedRow.Cells
				cell.RemoveAllChildren()
			Next cell

			' Add the row to the end of the table.
			table.AppendChild(clonedRow)

			doc.Save(MyDir & "\Artifacts\Table.AddCloneRowToTable.doc")
			'ExEnd

			' Verify that the row was cloned and appended properly.
			Assert.AreEqual(5, table.Rows.Count)
			Assert.AreEqual(String.Empty, table.LastRow.ToString(SaveFormat.Text).Trim())
			Assert.AreEqual(2, table.LastRow.Cells.Count)
		End Sub

		<Test> _
		Public Sub FixDefaultTableWidthsInAw105()
			'ExStart
			'ExId:FixTablesDefaultFixedColumnWidth
			'ExSummary:Shows how to revert the default behaviour of table sizing to use column widths.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Keep a reference to the table being built.
			Dim table As Table = builder.StartTable()

			' Apply some formatting.
			builder.CellFormat.Width = 100
			builder.CellFormat.Shading.BackgroundPatternColor = Color.Red

			builder.InsertCell()
			' This will cause the table to be structured using column widths as in previous verisons
			' instead of fitted to the page width like in the newer versions.
			table.AutoFit(AutoFitBehavior.FixedColumnWidths)

			' Continue with building your table as usual...
			'ExEnd
		End Sub

		<Test> _
		Public Sub FixDefaultTableBordersIn105()
			'ExStart
			'ExId:FixTablesDefaultBorders
			'ExSummary:Shows how to revert the default borders on tables back to no border lines.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Keep a reference to the table being built.
			Dim table As Table = builder.StartTable()

			builder.InsertCell()
			' Clear all borders to match the defaults used in previous versions.
			table.ClearBorders()

			' Continue with building your table as usual...
			'ExEnd
		End Sub

		<Test> _
		Public Sub FixDefaultTableFormattingExceptionIn105()
			'ExStart
			'ExId:FixTableFormattingException
			'ExSummary:Shows how to avoid encountering an exception when applying table formatting.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Keep a reference to the table being built.
			Dim table As Table = builder.StartTable()

			' We must first insert a new cell which in turn inserts a row into the table.
			builder.InsertCell()
			' Once a row exists in our table we can apply table wide formatting.
			table.AllowAutoFit = True

			' Continue with building your table as usual...
			'ExEnd
		End Sub

		<Test> _
		Public Sub FixRowFormattingNotAppliedIn105()
			'ExStart
			'ExId:FixRowFormattingNotApplied
			'ExSummary:Shows how to fix row formatting not being applied to some rows.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			builder.StartTable()

			' For the first row this will be set correctly.
			builder.RowFormat.HeadingFormat = True

			builder.InsertCell()
			builder.Writeln("Text")
			builder.InsertCell()
			builder.Writeln("Text")

			' End the first row.
			builder.EndRow()

			' Here we would normally define some other row formatting, such as disabling the 
			' heading format. However at the moment this will be ignored and the value from the 
			' first row reapplied to the row.

			builder.InsertCell()

			' Instead make sure to specify the row formatting for the second row here.
			builder.RowFormat.HeadingFormat = False

			' Continue with building your table as usual...
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetIndexOfTableElements()
			Dim doc As New Document(MyDir & "Table.Document.doc")

			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)
			'ExStart
			'ExFor:NodeCollection.IndexOf
			'ExId:IndexOfTable
			'ExSummary:Retrieves the index of a table in the document.
			Dim allTables As NodeCollection = doc.GetChildNodes(NodeType.Table, True)
			Dim tableIndex As Integer = allTables.IndexOf(table)
			'ExEnd

			Dim row As Row = table.Rows(2)
			'ExStart
			'ExFor:Row
			'ExFor:CompositeNode.IndexOf
			'ExId:IndexOfRow
			'ExSummary:Retrieves the index of a row in a table.
			Dim rowIndex As Integer = table.IndexOf(row)
			'ExEnd

			Dim cell As Cell = row.LastCell
			'ExStart
			'ExFor:Cell
			'ExFor:CompositeNode.IndexOf
			'ExId:IndexOfCell
			'ExSummary:Retrieves the index of a cell in a row.
			Dim cellIndex As Integer = row.IndexOf(cell)
			'ExEnd

			Assert.AreEqual(0, tableIndex)
			Assert.AreEqual(2, rowIndex)
			Assert.AreEqual(4, cellIndex)
		End Sub

		<Test> _
		Public Sub GetPreferredWidthTypeAndValue()
			Dim doc As New Document(MyDir & "Table.Document.doc")

			' Find the first table in the document
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)
			'ExStart
			'ExFor:PreferredWidthType
			'ExFor:PreferredWidth.Type
			'ExFor:PreferredWidth.Value
			'ExId:GetPreferredWidthTypeAndValue
			'ExSummary:Retrieves the preferred width type of a table cell.
			Dim firstCell As Cell = table.FirstRow.FirstCell
			Dim type As PreferredWidthType = firstCell.CellFormat.PreferredWidth.Type
			Dim value As Double = firstCell.CellFormat.PreferredWidth.Value
			'ExEnd

			Assert.AreEqual(PreferredWidthType.Percent, type)
			Assert.AreEqual(11.16, value)
		End Sub

		<Test> _
		Public Sub InsertTableUsingNodeConstructors()
			'ExStart
			'ExFor:Table
			'ExFor:Row
			'ExFor:Row.RowFormat
			'ExFor:RowFormat
			'ExFor:Cell
			'ExFor:Cell.CellFormat
			'ExFor:CellFormat
			'ExFor:CellFormat.Shading
			'ExFor:Cell.FirstParagraph
			'ExId:InsertTableUsingNodeConstructors
			'ExSummary:Shows how to insert a table using the constructors of nodes.
			Dim doc As New Document()

			' We start by creating the table object. Note how we must pass the document object
			' to the constructor of each node. This is because every node we create must belong
			' to some document.
			Dim table As New Table(doc)
			' Add the table to the document.
			doc.FirstSection.Body.AppendChild(table)

			' Here we could call EnsureMinimum to create the rows and cells for us. This method is used
			' to ensure that the specified node is valid, in this case a valid table should have at least one
			' row and one cell, therefore this method creates them for us.

			' Instead we will handle creating the row and table ourselves. This would be the best way to do this
			' if we were creating a table inside an algorthim for example.
			Dim row As New Row(doc)
			row.RowFormat.AllowBreakAcrossPages = True
			table.AppendChild(row)

			' We can now apply any auto fit settings.
			table.AutoFit(AutoFitBehavior.FixedColumnWidths)

			' Create a cell and add it to the row
			Dim cell As New Cell(doc)
			cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue
			cell.CellFormat.Width = 80

			' Add a paragraph to the cell as well as a new run with some text.
			cell.AppendChild(New Paragraph(doc))
			cell.FirstParagraph.AppendChild(New Run(doc, "Row 1, Cell 1 Text"))

			' Add the cell to the row.
			row.AppendChild(cell)

			' We would then repeat the process for the other cells and rows in the table.
			' We can also speed things up by cloning existing cells and rows.
			row.AppendChild(cell.Clone(False))
			row.LastCell.AppendChild(New Paragraph(doc))
			row.LastCell.FirstParagraph.AppendChild(New Run(doc, "Row 1, Cell 2 Text"))

			doc.Save(MyDir & "\Artifacts\Table.InsertTableUsingNodes.doc")
			'ExEnd

			Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, True).Count)
			Assert.AreEqual(1, doc.GetChildNodes(NodeType.Row, True).Count)
			Assert.AreEqual(2, doc.GetChildNodes(NodeType.Cell, True).Count)
			Assert.AreEqual("Row 1, Cell 1 Text" & Constants.vbCrLf & "Row 1, Cell 2 Text", doc.FirstSection.Body.Tables(0).ToString(SaveFormat.Text).Trim())
		End Sub

		'ExStart
		'ExFor:Table
		'ExFor:Row
		'ExFor:Cell
		'ExFor:Table.#ctor(DocumentBase)
		'ExFor:Row.#ctor(DocumentBase)
		'ExFor:Cell.#ctor(DocumentBase)
		'ExId:NestedTableNodeConstructors
		'ExSummary:Shows how to build a nested table without using DocumentBuilder.
		<Test> _
		Public Sub NestedTablesUsingNodeConstructors()
			Dim doc As New Document()

			' Create the outer table with three rows and four columns.
			Dim outerTable As Table = Me.CreateTable(doc, 3, 4, "Outer Table")
			' Add it to the document body.
			doc.FirstSection.Body.AppendChild(outerTable)

			' Create another table with two rows and two columns.
			Dim innerTable As Table = Me.CreateTable(doc, 2, 2, "Inner Table")
			' Add this table to the first cell of the outer table.
			outerTable.FirstRow.FirstCell.AppendChild(innerTable)

			doc.Save(MyDir & "\Artifacts\Table.CreateNestedTable.doc")

			Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, True).Count) ' ExSkip
			Assert.AreEqual(1, outerTable.FirstRow.FirstCell.Tables.Count) 'ExSkip
			Assert.AreEqual(16, outerTable.GetChildNodes(NodeType.Cell, True).Count) 'ExSkip
			Assert.AreEqual(4, innerTable.GetChildNodes(NodeType.Cell, True).Count) 'ExSkip
		End Sub

		''' <summary>
		''' Creates a new table in the document with the given dimensions and text in each cell.
		''' </summary>
		Private Function CreateTable(ByVal doc As Document, ByVal rowCount As Integer, ByVal cellCount As Integer, ByVal cellText As String) As Table
			Dim table As New Table(doc)

			' Create the specified number of rows.
			For rowId As Integer = 1 To rowCount
				Dim row As New Row(doc)
				table.AppendChild(row)

				' Create the specified number of cells for each row.
				For cellId As Integer = 1 To cellCount
					Dim cell As New Cell(doc)
					row.AppendChild(cell)
					' Add a blank paragraph to the cell.
					cell.AppendChild(New Paragraph(doc))

					' Add the text.
					cell.FirstParagraph.AppendChild(New Run(doc, cellText))
				Next cellId
			Next rowId

			Return table
		End Function
		'ExEnd

		'ExStart
		'ExFor:CellFormat.HorizontalMerge
		'ExFor:CellFormat.VerticalMerge
		'ExFor:CellMerge
		'ExId:CheckCellMerge
		'ExSummary:Prints the horizontal and vertical merge type of a cell.
		<Test> _
		Public Sub CheckCellsMerged()
			Dim doc As New Document(MyDir & "Table.MergedCells.doc")

			' Retrieve the first table in the document.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			For Each row As Row In table.Rows
				For Each cell As Cell In row.Cells
					Console.WriteLine(Me.PrintCellMergeType(cell))
				Next cell
			Next row

			Assert.AreEqual("The cell at R1, C1 is horizontally merged.", Me.PrintCellMergeType(table.FirstRow.FirstCell)) 'ExSkip
		End Sub

		Public Function PrintCellMergeType(ByVal cell As Cell) As String
            Dim isHorizontallyMerged As Boolean = cell.CellFormat.HorizontalMerge <> CellMerge.None
            Dim isVerticallyMerged As Boolean = cell.CellFormat.VerticalMerge <> CellMerge.None
			Dim cellLocation As String = String.Format("R{0}, C{1}", cell.ParentRow.ParentTable.IndexOf(cell.ParentRow) + 1, cell.ParentRow.IndexOf(cell) + 1)

			If isHorizontallyMerged AndAlso isVerticallyMerged Then
				Return String.Format("The cell at {0} is both horizontally and vertically merged", cellLocation)
			ElseIf isHorizontallyMerged Then
				Return String.Format("The cell at {0} is horizontally merged.", cellLocation)
			ElseIf isVerticallyMerged Then
				Return String.Format("The cell at {0} is vertically merged", cellLocation)
			Else
				Return String.Format("The cell at {0} is not merged", cellLocation)
			End If
		End Function
		'ExEnd

		<Test> _
		Public Sub MergeCellRange()
			' Open the document
			Dim doc As New Document(MyDir & "Table.Document.doc")

			' Retrieve the first table in the body of the first section.
			Dim table As Table = doc.FirstSection.Body.Tables(0)

			'ExStart
			'ExId:MergeCellRange
			'ExSummary:Merges the range of cells between the two specified cells.
			' We want to merge the range of cells found inbetween these two cells.
			Dim cellStartRange As Cell = table.Rows(2).Cells(2)
			Dim cellEndRange As Cell = table.Rows(3).Cells(3)

			' Merge all the cells between the two specified cells into one.
			MergeCells(cellStartRange, cellEndRange)
			'ExEnd

			' Save the document.
			doc.Save(MyDir & "\Artifacts\Table.MergeCellRange.doc")

			' Verify the cells were merged
			Dim mergedCellsCount As Integer = 0
			For Each cell As Cell In table.GetChildNodes(NodeType.Cell, True)
                If cell.CellFormat.HorizontalMerge <> CellMerge.None OrElse cell.CellFormat.HorizontalMerge <> CellMerge.None Then
                    mergedCellsCount += 1
                End If
			Next cell

			Assert.AreEqual(4, mergedCellsCount)
            Assert.True(table.Rows(2).Cells(2).CellFormat.HorizontalMerge = CellMerge.First)
            Assert.True(table.Rows(2).Cells(2).CellFormat.VerticalMerge = CellMerge.First)
            Assert.True(table.Rows(3).Cells(3).CellFormat.HorizontalMerge = CellMerge.Previous)
            Assert.True(table.Rows(3).Cells(3).CellFormat.VerticalMerge = CellMerge.Previous)
		End Sub

		'ExStart
		'ExId:MergeCellsMethod
		'ExSummary:A method which merges all cells of a table in the specified range of cells.
		''' <summary>
		''' Merges the range of cells found between the two specified cells both horizontally and vertically. Can span over multiple rows.
		''' </summary>
		Public Shared Sub MergeCells(ByVal startCell As Cell, ByVal endCell As Cell)
			Dim parentTable As Table = startCell.ParentRow.ParentTable

			' Find the row and cell indices for the start and end cell.
			Dim startCellPos As New Point(startCell.ParentRow.IndexOf(startCell), parentTable.IndexOf(startCell.ParentRow))
			Dim endCellPos As New Point(endCell.ParentRow.IndexOf(endCell), parentTable.IndexOf(endCell.ParentRow))
			' Create the range of cells to be merged based off these indices. Inverse each index if the end cell if before the start cell. 
			Dim mergeRange As New Rectangle(Math.Min(startCellPos.X, endCellPos.X), Math.Min(startCellPos.Y, endCellPos.Y), Math.Abs(endCellPos.X - startCellPos.X) + 1, Math.Abs(endCellPos.Y - startCellPos.Y) + 1)

			For Each row As Row In parentTable.Rows
				For Each cell As Cell In row.Cells
					Dim currentPos As New Point(row.IndexOf(cell), parentTable.IndexOf(row))

					' Check if the current cell is inside our merge range then merge it.
					If mergeRange.Contains(currentPos) Then
						If currentPos.X = mergeRange.X Then
							cell.CellFormat.HorizontalMerge = CellMerge.First
						Else
							cell.CellFormat.HorizontalMerge = CellMerge.Previous
						End If

						If currentPos.Y = mergeRange.Y Then
							cell.CellFormat.VerticalMerge = CellMerge.First
						Else
							cell.CellFormat.VerticalMerge = CellMerge.Previous
						End If
					End If
				Next cell
			Next row
		End Sub
		'ExEnd

		<Test> _
		Public Sub CombineTables()
			'ExStart
			'ExFor:Table
			'ExFor:Cell.CellFormat
			'ExFor:CellFormat.Borders
			'ExFor:Table.Rows
			'ExFor:Table.FirstRow
			'ExFor:CellFormat.ClearFormatting
			'ExId:CombineTables
			'ExSummary:Shows how to combine the rows from two tables into one.
			' Load the document.
			Dim doc As New Document(MyDir & "Table.Document.doc")

			' Get the first and second table in the document.
			' The rows from the second table will be appended to the end of the first table.
			Dim firstTable As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)
			Dim secondTable As Table = CType(doc.GetChild(NodeType.Table, 1, True), Table)

			' Append all rows from the current table to the next.
			' Due to the design of tables even tables with different cell count and widths can be joined into one table.
			Do While secondTable.HasChildNodes
				firstTable.Rows.Add(secondTable.FirstRow)
			Loop

			' Remove the empty table container.
			secondTable.Remove()

			doc.Save(MyDir & "\Artifacts\Table.CombineTables.doc")
			'ExEnd

			Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, True).Count)
			Assert.AreEqual(9, doc.FirstSection.Body.Tables(0).Rows.Count)
			Assert.AreEqual(42, doc.FirstSection.Body.Tables(0).GetChildNodes(NodeType.Cell, True).Count)
		End Sub

		<Test> _
		Public Sub SplitTable()
			'ExStart
			'ExId:SplitTableAtRow
			'ExSummary:Shows how to split a table into two tables a specific row.
			' Load the document.
			Dim doc As New Document(MyDir & "Table.SimpleTable.doc")

			' Get the first table in the document.
			Dim firstTable As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

			' We will split the table at the third row (inclusive).
			Dim row As Row = firstTable.Rows(2)

			' Create a new container for the split table.
			Dim table As Table = CType(firstTable.Clone(False), Table)

			' Insert the container after the original.
			firstTable.ParentNode.InsertAfter(table, firstTable)

			' Add a buffer paragraph to ensure the tables stay apart.
			firstTable.ParentNode.InsertAfter(New Paragraph(doc), firstTable)

			Dim currentRow As Row

			Do
				currentRow = firstTable.LastRow
				table.PrependChild(currentRow)
			Loop While currentRow IsNot row

			doc.Save(MyDir & "\Artifacts\Table.SplitTable.doc")
			'ExEnd

			doc = New Document(MyDir & "\Artifacts\Table.SplitTable.doc")
			' Test we are adding the rows in the correct order and the 
			' selected row was also moved.
			Assert.AreEqual(row, table.FirstRow)

			Assert.AreEqual(2, firstTable.Rows.Count)
			Assert.AreEqual(2, table.Rows.Count)
			Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, True).Count)
		End Sub
	End Class
End Namespace
