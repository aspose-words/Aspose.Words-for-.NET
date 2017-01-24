' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports Aspose.Words.Tables

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExCellFormat
		Inherits ApiExampleBase
		<Test> _
		Public Sub VerticalMerge()
			'ExStart
			'ExFor:DocumentBuilder.InsertCell
			'ExFor:DocumentBuilder.EndRow
			'ExFor:CellMerge
			'ExFor:CellFormat.VerticalMerge
			'ExId:VerticalMerge
			'ExSummary:Creates a table with two columns with cells merged vertically in the first column.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			builder.InsertCell()
			builder.CellFormat.VerticalMerge = CellMerge.First
			builder.Write("Text in merged cells.")

			builder.InsertCell()
			builder.CellFormat.VerticalMerge = CellMerge.None
			builder.Write("Text in one cell")
			builder.EndRow()

			builder.InsertCell()
			' This cell is vertically merged to the cell above and should be empty.
			builder.CellFormat.VerticalMerge = CellMerge.Previous

			builder.InsertCell()
			builder.CellFormat.VerticalMerge = CellMerge.None
			builder.Write("Text in another cell")
			builder.EndRow()
			builder.EndTable()
			'ExEnd
		End Sub

		<Test> _
		Public Sub HorizontalMerge()
			'ExStart
			'ExFor:CellMerge
			'ExFor:CellFormat.HorizontalMerge
			'ExId:HorizontalMerge
			'ExSummary:Creates a table with two rows with cells in the first row horizontally merged.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			builder.InsertCell()
			builder.CellFormat.HorizontalMerge = CellMerge.First
			builder.Write("Text in merged cells.")

			builder.InsertCell()
			' This cell is merged to the previous and should be empty.
			builder.CellFormat.HorizontalMerge = CellMerge.Previous
			builder.EndRow()

			builder.InsertCell()
			builder.CellFormat.HorizontalMerge = CellMerge.None
			builder.Write("Text in one cell.")

			builder.InsertCell()
			builder.Write("Text in another cell.")
			builder.EndRow()
			builder.EndTable()
			'ExEnd
		End Sub
	End Class
End Namespace
