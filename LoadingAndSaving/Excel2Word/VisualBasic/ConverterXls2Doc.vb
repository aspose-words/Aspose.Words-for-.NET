' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.Drawing
Imports Aspose.Cells
Imports Aspose.Cells.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables


Namespace Excel2Word
	''' <summary>
	''' This class contains methods that allow to convert Excel workbook to Word document.
	''' 
	''' A main idea is that Excel table can have unlimited witdh, 
	''' that's why we should split this table, how Excel do in print preview.
	''' 
	''' This class demonstrates how you can use Aspose libraries to convert different types of document.
	''' </summary>
	Friend Class ConverterXls2Doc
		''' <summary>
		''' Convert Excel workbook to Word document
		''' </summary>
		''' <param name="workbook">Input workbook</param>
		''' <returns>Word document</returns>
		Friend Function Convert(ByVal workbook As Workbook) As Document
			'Create new document
			Dim doc As New Document()
			'Create an instance of the  DocumentBuilder class
			Dim builder As New DocumentBuilder(doc)

			'Every worksheet in Excel workbook is represented as section in Word document
			For Each worksheet As Worksheet In workbook.Worksheets
				'Import PageSetup from Excel file to Word document
				'Orientation can be Portrait or Landscape
				builder.PageSetup.Orientation = ConvertPageOrientation(worksheet.PageSetup.Orientation)
				'Paper size can be A4, A3, Letter, etc.
				builder.PageSetup.PaperSize = ConvertPaperSize(worksheet.PageSetup.PaperSize)
				'Import margins
				builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(worksheet.PageSetup.LeftMarginInch) ' 1cm = 28.35pt
				builder.PageSetup.RightMargin = ConvertUtil.InchToPoint(worksheet.PageSetup.RightMarginInch)
				builder.PageSetup.TopMargin = ConvertUtil.InchToPoint(worksheet.PageSetup.TopMarginInch)
				builder.PageSetup.BottomMargin = ConvertUtil.InchToPoint(worksheet.PageSetup.BottomMarginInch)

				'Get array of Word tables, every table in this array represents a part of Excel worksheet.
				Dim partsArray As ArrayList = GetTablePartsArray(worksheet, doc)
				'Insert all tables into the Word document
				For Each table As Table In partsArray
					'Insert table
					builder.CurrentSection.Body.AppendChild(table)
					'Move coursore to document end
					builder.MoveToDocumentEnd()
					'Insert break if table is not last in the collection
					If (Not table.Equals(partsArray(partsArray.Count - 1))) Then
						builder.InsertBreak(BreakType.SectionBreakNewPage)
					End If
				Next table
				'Insert break if current workseet is not last in the Excwl workbook
				If (Not worksheet.Equals(workbook.Worksheets(workbook.Worksheets.Count - 1))) AndAlso partsArray.Count <> 0 Then
					builder.InsertBreak(BreakType.SectionBreakNewPage)
				End If
			Next worksheet

			Return doc
		End Function

		''' <summary>
		''' This method returns array of Word tables, every table in this array represents a part of Excel worksheet.
		''' </summary>
		''' <param name="excelWorksheet">Input worksheet</param>
		''' <param name="doc">Parent document</param>
		''' <returns>Array of Word tables</returns>
		Private Function GetTablePartsArray(ByVal excelWorksheet As Worksheet, ByVal doc As Document) As ArrayList
			'Get column index of cell that contains data
			Dim colCount As Integer = excelWorksheet.Cells.MaxColumn + 1
			'Get row index of cell that contains data
			Dim rowCount As Integer = excelWorksheet.Cells.MaxRow + 1
			Dim startColumn As Integer = excelWorksheet.Cells.MinColumn
			Dim startRow As Integer = excelWorksheet.Cells.MinRow

			'Get area in the worksheet that will be printed
			'Returns something like this "A1:D51" of null
			Dim excelPrintArea As String = excelWorksheet.PageSetup.PrintArea
			If (Not String.IsNullOrEmpty(excelPrintArea)) Then
				'Get first cell in the printed area
				Dim rangeStart As String = excelPrintArea.Substring(0, excelPrintArea.IndexOf(":"))
				'Get last cell in the printed area
				Dim rangeEnd As String = excelPrintArea.Substring(excelPrintArea.IndexOf(":") + 1)
				'Get printed range from worksheet
				Dim range As Aspose.Cells.Range = excelWorksheet.Cells.CreateRange(rangeStart, rangeEnd)

				colCount = range.ColumnCount + range.FirstColumn
				rowCount = range.RowCount + range.FirstRow
				startColumn = range.FirstColumn
				startRow = range.FirstRow
			End If

			'Extract objects like Pictures, Charts, etc and store in the HashTable
			'if worksheet contains object that is placed outside the region then resize region (count of rows and columns)
			Dim drawRange As Aspose.Cells.Range = ExtractDrawingObjects(excelWorksheet)
			If drawRange.RowCount > rowCount Then
				rowCount = drawRange.RowCount
			End If
			If drawRange.ColumnCount > colCount Then
				colCount = drawRange.ColumnCount
			End If
			If drawRange.FirstRow < startRow Then
				startRow = drawRange.FirstRow
			End If
			If drawRange.FirstColumn < startColumn Then
				startColumn = drawRange.FirstColumn
			End If

			'Create ampty ArrayList
			Dim tablePartList As New ArrayList()
			'split worksheet into a tables collection
			GetTablePart(excelWorksheet, doc, tablePartList, startColumn, startRow, colCount, rowCount)

			Return tablePartList
		End Function

		''' <summary>
		''' This Method splits worksheet into a tables collection.
		''' </summary>
		''' <param name="excelWorksheet">Input Worksheet</param>
		''' <param name="doc">Parent document</param>
		''' <param name="tablePartList">ArrayList where tables will be stored</param>
		''' <param name="columnStartIndex">Index of the column in Excel worksheet that will be the first column of Word table</param>
		''' <param name="rowStartIndex">Index of the row in Excel worksheet that will be the first row of Word table</param>
		''' <param name="columnCount">Column index of a last cell that contains data in the Excel worksheet</param>
		''' <param name="rowCount">Row index of a last cell that contains data in the Excel worksheet</param>
		Private Sub GetTablePart(ByVal excelWorksheet As Worksheet, ByVal doc As Document, ByVal tablePartList As ArrayList, ByVal columnStartIndex As Integer, ByVal rowStartIndex As Integer, ByVal columnCount As Integer, ByVal rowCount As Integer)
			If columnCount <> 0 AndAlso rowCount <> 0 Then
				'Calculate max width of Words table. 
				'Then we will add columns to Word table while it's width is < maxWidth
				Dim setup As Aspose.Words.PageSetup = doc.FirstSection.PageSetup
				Dim maxWidth As Double = setup.PageWidth - setup.LeftMargin - setup.RightMargin
				Dim newColumnStartIndex As Integer = 0
				Dim currentWidth As Double = 0
				For columnIndex As Integer = columnStartIndex To columnCount
					'Calculate width of current Word table
					currentWidth += ConvertUtil.PixelToPoint(excelWorksheet.Cells.GetColumnWidthPixel(columnIndex))
					newColumnStartIndex = columnIndex
					'If width of table > maxWidth then break loop
					If currentWidth > maxWidth AndAlso columnIndex <> columnStartIndex Then
						Exit For
					End If
				Next columnIndex
				'Create a new Word table
				Dim wordsTable As New Table(doc)
				'Loop through rows in the Excel worksheet
				For rowIndex As Integer = rowStartIndex To rowCount - 1
					'Create new row
					Dim wordsRow As New Aspose.Words.Tables.Row(doc)
					'Get cllection of Excel cells
					Dim cells As Aspose.Cells.Cells = excelWorksheet.Cells
					'Set height of current row
					wordsRow.RowFormat.Height = ConvertUtil.PixelToPoint(cells.GetRowHeightPixel(rowIndex))
					'Append current row to current table.
					wordsTable.AppendChild(wordsRow)
					'Loop through columns and add columns to Word table while table's width < maxWidth
					For columnIndex As Integer = columnStartIndex To newColumnStartIndex - 1
						'Convert Excel cell to Word cell
						Dim wordsCell As Aspose.Words.Tables.Cell = ImportExcelCell(doc, cells, rowIndex, columnIndex)
						'Insert cell into rhe row
						wordsRow.AppendChild(wordsCell)
					Next columnIndex
				Next rowIndex

				' We want the table to take only as much of the page as required.
				wordsTable.PreferredWidth = PreferredWidth.Auto

				'Add Word table to ArrayList
				tablePartList.Add(wordsTable)

				If newColumnStartIndex < columnCount Then
					'Start next table from newColumnStartIndex
					GetTablePart(excelWorksheet, doc, tablePartList, newColumnStartIndex, rowStartIndex, columnCount, rowCount)
				End If
			End If
		End Sub

		''' <summary>
		''' Convert Excel Cell to Word Cell
		''' </summary>
		''' <param name="doc">Parent document</param>
		''' <param name="cells">Excel cells collection</param>
		''' <param name="rowIndex">Row index</param>
		''' <param name="columnIndex">Column index</param>
		''' <returns>Word Cell</returns>
		Private Function ImportExcelCell(ByVal doc As Document, ByVal cells As Aspose.Cells.Cells, ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Aspose.Words.Tables.Cell
			'Create a new Word Cell
			Dim wordsCell As New Aspose.Words.Tables.Cell(doc)
			'Get Excel cell from collection
			Dim excelCell As Aspose.Cells.Cell = cells(rowIndex, columnIndex)
			'Set cell width
			Dim cellWidth As Double = ConvertUtil.PixelToPoint(cells.GetColumnWidthPixel(columnIndex))
			wordsCell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(cellWidth)
			wordsCell.CellFormat.Width = ConvertUtil.PixelToPoint(cellWidth)
			'Set background color
			wordsCell.CellFormat.Shading.ForegroundPatternColor = excelCell.GetDisplayStyle().ForegroundColor
			wordsCell.CellFormat.Shading.BackgroundPatternColor = excelCell.GetDisplayStyle().BackgroundColor
			'Set background texture
			wordsCell.CellFormat.Shading.Texture = ConvertBackgroundTexture(excelCell.GetDisplayStyle().Pattern)
			'Import borders from Excel cell to Word cell
			ImportBorders(wordsCell, excelCell)
			'Set vertical alignment
			wordsCell.CellFormat.VerticalAlignment = ConvertVerticalAlignment(excelCell.GetDisplayStyle().VerticalAlignment)
			'If Excel cells is merged then merge cells in Word Table
			wordsCell.CellFormat.VerticalMerge = ConvertVerticalCellMerge(excelCell)
			wordsCell.CellFormat.HorizontalMerge = ConvertHorizontalCellMerge(excelCell)
			'Create paragraph that will containc content of cell
			Dim wordsParagraph As New Paragraph(doc)
			'Set horizontal alignment
			wordsParagraph.ParagraphFormat.Alignment = ConvertHorizontalAlignment(excelCell.GetDisplayStyle().HorizontalAlignment)
			'Get text with formating from Excel cell as collection Run
			Dim wordRuns As ArrayList = GetTextFromCell(excelCell, doc)
			For Each run As Run In wordRuns
				wordsParagraph.AppendChild(run)
			Next run
			'Import formating of the cell
			ImportFont(wordsParagraph.ParagraphBreakFont, excelCell.GetDisplayStyle().Font)
			'Insert paragrahp with content into cell
			wordsCell.AppendChild(wordsParagraph)
			'If Excel cell contains drawing object then convert this object and insert into Word cell
			InsertDrawingObject(excelCell, wordsCell)

			Return wordsCell
		End Function

		''' <summary>
		''' Inserts excel drawing object into Word cell
		''' </summary>
		''' <param name="excelCell">Excel cell</param>
		''' <param name="wordsCell">Word cell</param>
		Private Sub InsertDrawingObject(ByVal excelCell As Aspose.Cells.Cell, ByVal wordsCell As Aspose.Words.Tables.Cell)
			'If current cell is horizontaly merged with previose cell then we should calculate offset of Shape
			Dim leftOffset As Double = GetAdditionalHorizontalOffset(excelCell)
			'If current cell is verticaly merged with previose cell then we should calculate offset of Shape
			Dim topOffset As Double = GetAdditionalVerticalOffset(excelCell)

			If mPicturesCollection.ContainsKey(excelCell) Then
				'Get Picture object from HashTable
				Dim excelPicture As Aspose.Cells.Drawing.Picture = CType(mPicturesCollection(excelCell), Aspose.Cells.Drawing.Picture)
				'Convert Excel Picture to Word Shape
				Dim wordsShape As Aspose.Words.Drawing.Shape = ConvertPictureToShape(excelPicture, wordsCell.Document)
				wordsShape.Left += leftOffset
				wordsShape.Top += topOffset
				'Insert Shape into current cell
				wordsCell.LastParagraph.AppendChild(wordsShape)
			End If
			If mChartsCollection.ContainsKey(excelCell) Then
				'Get Chart object from HashTable
				Dim excelChart As Aspose.Cells.Charts.Chart = CType(mChartsCollection(excelCell), Aspose.Cells.Charts.Chart)
				'Convert Excel Chart to Word Shape
				Dim wordsShape As Aspose.Words.Drawing.Shape = ConvertCartToShape(excelChart, wordsCell.Document)
				If wordsShape IsNot Nothing Then
					wordsShape.Left += leftOffset
					wordsShape.Top += topOffset
					wordsCell.LastParagraph.AppendChild(wordsShape)
				End If
			End If
			If mCheckBoxesCollection.ContainsKey(excelCell) Then
				'Get CheckBox object from HashTable
				Dim excelCheckBox As Aspose.Cells.Drawing.CheckBox = CType(mCheckBoxesCollection(excelCell), Aspose.Cells.Drawing.CheckBox)
				'Insert CheckBox into the current cell
				InsertCheckBox(excelCheckBox, wordsCell)
			End If
			If mTextBoxesCollection.ContainsKey(excelCell) Then
				'Get TextBox object from HashTable
				Dim excelTextBox As Aspose.Cells.Drawing.TextBox = CType(mTextBoxesCollection(excelCell), Aspose.Cells.Drawing.TextBox)
				'Convert Excel TextBox to Word TextBox
				Dim textBox As Aspose.Words.Drawing.Shape = ConvertTextBoxToShape(excelTextBox, wordsCell.Document)
				textBox.Left += leftOffset
				textBox.Top += topOffset
				'Insert Shape into current cell
				wordsCell.LastParagraph.AppendChild(textBox)
			End If
			If mShapesCollection.ContainsKey(excelCell) Then
				'Get Shape object from HashTable
				Dim excelShape As Shape = CType(mShapesCollection(excelCell), Shape)
				'Convert Excel Shape to Word Shape
				Dim wordsShape As Aspose.Words.Drawing.Shape = ConvertShapeToShape(excelShape, wordsCell.Document)
				wordsShape.Left += leftOffset
				wordsShape.Top += topOffset
				'Insert Shape into current cell
				wordsCell.LastParagraph.AppendChild(wordsShape)
			End If
		End Sub

		''' <summary>
		''' Extract objects like Pictures, Charts, etc and store in the HashTable
		''' </summary>
		''' <param name="excelWorksheet">Excel worksheet</param>
		''' <returns>Range that contains drawing object</returns>
		Private Function ExtractDrawingObjects(ByVal excelWorksheet As Worksheet) As Aspose.Cells.Range
			Dim rowIndex As Integer = 0
			Dim columnIndex As Integer = 0
			Dim lastRow As Integer = 0
			Dim lastColumn As Integer = 0
			Dim firstRow As Integer = excelWorksheet.Cells.Rows.Count
			Dim firstColumn As Integer = excelWorksheet.Cells.Columns.Count

			'Get collection of Pictures in current worksheet
			Dim pictures As Aspose.Cells.Drawing.PictureCollection = excelWorksheet.Pictures
			For Each picture As Aspose.Cells.Drawing.Picture In pictures
				rowIndex = picture.UpperLeftRow
				columnIndex = picture.UpperLeftColumn
				If columnIndex > lastColumn Then
					lastColumn = columnIndex
				End If
				If rowIndex > lastRow Then
					lastRow = rowIndex
				End If
				If columnIndex < firstColumn Then
					firstColumn = columnIndex
				End If
				If rowIndex < firstRow Then
					firstRow = rowIndex
				End If
				'Add Picture to HashTable. Key is upper left cell and value is Picture
				mPicturesCollection.Add(excelWorksheet.Cells(rowIndex, columnIndex), picture)
			Next picture
			'Get collection of Charts in current worksheet
			Dim charts As Aspose.Cells.Charts.ChartCollection = excelWorksheet.Charts
			For Each chart As Aspose.Cells.Charts.Chart In charts
				rowIndex = chart.ChartObject.UpperLeftRow
				columnIndex = chart.ChartObject.UpperLeftColumn
				If columnIndex > lastColumn Then
					lastColumn = columnIndex
				End If
				If rowIndex > lastRow Then
					lastRow = rowIndex
				End If
				If columnIndex < firstColumn Then
					firstColumn = columnIndex
				End If
				If rowIndex < firstRow Then
					firstRow = rowIndex
				End If
				'Add Chart to HashTable. Key is upper left cell and value is Chart
				mChartsCollection.Add(excelWorksheet.Cells(rowIndex, columnIndex), chart)
			Next chart
			'Get collection of CheckBoxes in current worksheet
			Dim checkBoxes As Aspose.Cells.Drawing.CheckBoxCollection = excelWorksheet.CheckBoxes
			For Each checkBox As Aspose.Cells.Drawing.CheckBox In checkBoxes
				rowIndex = checkBox.UpperLeftRow
				columnIndex = checkBox.UpperLeftColumn
				If columnIndex > lastColumn Then
					lastColumn = columnIndex
				End If
				If rowIndex > lastRow Then
					lastRow = rowIndex
				End If
				If columnIndex < firstColumn Then
					firstColumn = columnIndex
				End If
				If rowIndex < firstRow Then
					firstRow = rowIndex
				End If
				'Add CheckBox to HashTable. Key is upper left cell and value is ChekBox
				mCheckBoxesCollection.Add(excelWorksheet.Cells(rowIndex, columnIndex), checkBox)
			Next checkBox
			'Get collection of TextBoxes in current worksheet
			Dim textBoxes As Aspose.Cells.Drawing.TextBoxCollection = excelWorksheet.TextBoxes
			For Each textBox As Aspose.Cells.Drawing.TextBox In textBoxes
				rowIndex = textBox.UpperLeftRow
				columnIndex = textBox.UpperLeftColumn
				If columnIndex > lastColumn Then
					lastColumn = columnIndex
				End If
				If rowIndex > lastRow Then
					lastRow = rowIndex
				End If
				If columnIndex < firstColumn Then
					firstColumn = columnIndex
				End If
				If rowIndex < firstRow Then
					firstRow = rowIndex
				End If
				'Add TextBox to HashTable. Key is upper left cell and value is ChekBox
				mTextBoxesCollection.Add(excelWorksheet.Cells(rowIndex, columnIndex), textBox)
			Next textBox
			'Get collection of Shapes
			Dim shapes As Aspose.Cells.Drawing.ShapeCollection = excelWorksheet.Shapes
			For Each shape As Shape In shapes
				If shape.MsoDrawingType = MsoDrawingType.Line OrElse shape.MsoDrawingType = MsoDrawingType.Arc OrElse shape.MsoDrawingType = MsoDrawingType.Oval OrElse shape.MsoDrawingType = MsoDrawingType.Rectangle Then
					rowIndex = shape.UpperLeftRow
					columnIndex = shape.UpperLeftColumn
					If columnIndex > lastColumn Then
						lastColumn = columnIndex
					End If
					If rowIndex > lastRow Then
						lastRow = rowIndex
					End If
					If columnIndex < firstColumn Then
						firstColumn = columnIndex
					End If
					If rowIndex < firstRow Then
						firstRow = rowIndex
					End If
					'Add TextBox to HashTable. Key is upper left cell and value is ChekBox
					mShapesCollection.Add(excelWorksheet.Cells(rowIndex, columnIndex), shape)
				End If
			Next shape

			Dim range As Aspose.Cells.Range = excelWorksheet.Cells.CreateRange(firstRow, firstColumn, lastRow + 1, lastColumn + 1)

			Return range
		End Function

		''' <summary>
		''' Import boreders from Excel cell to Word Cell
		''' </summary>
		''' <param name="wordsCell">Destination Word cell</param>
		''' <param name="excelCell">Source Excel cell</param>
		Private Sub ImportBorders(ByVal wordsCell As Aspose.Words.Tables.Cell, ByVal excelCell As Aspose.Cells.Cell)
			'Set line style
			wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Bottom).LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.BottomBorder).LineStyle)
			'Set line color
			If wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Bottom).LineStyle <> LineStyle.None Then
				wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Bottom).Color = excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.BottomBorder).Color
			End If

			wordsCell.CellFormat.Borders(Aspose.Words.BorderType.DiagonalDown).LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.DiagonalDown).LineStyle)
			If wordsCell.CellFormat.Borders(Aspose.Words.BorderType.DiagonalDown).LineStyle<> LineStyle.None Then
				wordsCell.CellFormat.Borders(Aspose.Words.BorderType.DiagonalDown).Color = excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.DiagonalDown).Color
			End If

			wordsCell.CellFormat.Borders(Aspose.Words.BorderType.DiagonalUp).LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.DiagonalUp).LineStyle)
			If wordsCell.CellFormat.Borders(Aspose.Words.BorderType.DiagonalUp).LineStyle<> LineStyle.None Then
				wordsCell.CellFormat.Borders(Aspose.Words.BorderType.DiagonalUp).Color = excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.DiagonalUp).Color
			End If

			wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Left).LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.LeftBorder).LineStyle)
			If wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Left).LineStyle <> LineStyle.None Then
				wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Left).Color = excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.LeftBorder).Color
			End If

			wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Right).LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.RightBorder).LineStyle)
			If wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Right).LineStyle <> LineStyle.None Then
				wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Right).Color = excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.RightBorder).Color
			End If

			wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Top).LineStyle = ConvertLineStyle(excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.TopBorder).LineStyle)
			If wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Top).LineStyle <> LineStyle.None Then
				wordsCell.CellFormat.Borders(Aspose.Words.BorderType.Top).Color = excelCell.GetDisplayStyle().Borders(Aspose.Cells.BorderType.TopBorder).Color
			End If
		End Sub

		''' <summary>
		''' Convert Excel BorderType to Word LineStyle
		''' </summary>
		''' <param name="lineType">Excel LineType</param>
		''' <returns>Word LineStyle</returns>
		Private Function ConvertLineStyle(ByVal lineType As CellBorderType) As LineStyle
			Dim wordsLineStyle As LineStyle = LineStyle.None
			Select Case lineType
				Case CellBorderType.DashDot
						wordsLineStyle = LineStyle.DotDash
						Exit Select
				Case CellBorderType.DashDotDot
						wordsLineStyle = LineStyle.DotDotDash
						Exit Select
				Case CellBorderType.Dashed
						wordsLineStyle = LineStyle.DashSmallGap
						Exit Select
				Case CellBorderType.Dotted
						wordsLineStyle = LineStyle.Dot
						Exit Select
				Case CellBorderType.Double
						wordsLineStyle = LineStyle.Double
						Exit Select
				Case CellBorderType.Hair
						wordsLineStyle = LineStyle.Hairline
						Exit Select
				Case CellBorderType.Medium
						wordsLineStyle = LineStyle.Single
						Exit Select
				Case CellBorderType.MediumDashDot
						wordsLineStyle = LineStyle.DotDash
						Exit Select
				Case CellBorderType.MediumDashDotDot
						wordsLineStyle = LineStyle.DotDotDash
						Exit Select
				Case CellBorderType.MediumDashed
						wordsLineStyle = LineStyle.DashSmallGap
						Exit Select
				Case CellBorderType.None
						wordsLineStyle = LineStyle.None
						Exit Select
				Case CellBorderType.SlantedDashDot
						wordsLineStyle = LineStyle.DotDash
						Exit Select
				Case CellBorderType.Thick
						wordsLineStyle = LineStyle.Thick
						Exit Select
				Case CellBorderType.Thin
						wordsLineStyle = LineStyle.Single
						Exit Select
				Case Else
						wordsLineStyle = LineStyle.None
						Exit Select
			End Select

			Return wordsLineStyle
		End Function

		''' <summary>
		''' Conver Excel Undeline to Word Underline
		''' </summary>
		''' <param name="underlineType">Excel UnderlineType</param>
		''' <returns>Word Underline</returns>
		Private Function ConvertUnderline(ByVal underlineType As FontUnderlineType) As Underline
			Dim wordsUnderline As Underline = Underline.None

			Select Case underlineType
				Case FontUnderlineType.Accounting
						wordsUnderline = Underline.Wavy
						Exit Select
				Case FontUnderlineType.Double
						wordsUnderline = Underline.Double
						Exit Select
				Case FontUnderlineType.DoubleAccounting
						wordsUnderline = Underline.WavyDouble
						Exit Select
				Case FontUnderlineType.None
						wordsUnderline = Underline.None
						Exit Select
				Case FontUnderlineType.Single
						wordsUnderline = Underline.Single
						Exit Select
				Case Else
						wordsUnderline = Underline.None
						Exit Select
			End Select

			Return wordsUnderline
		End Function

		''' <summary>
		''' Conver Excel VerticalAlignment to Word CellVerticalAlignment
		''' </summary>
		''' <param name="alignmentType">Excel AlignmentType</param>
		''' <returns>Word CellVerticalAlignment</returns>
		Private Function ConvertVerticalAlignment(ByVal alignmentType As TextAlignmentType) As CellVerticalAlignment
			Dim wordsAlignment As CellVerticalAlignment = CellVerticalAlignment.Top

			Select Case alignmentType
				Case TextAlignmentType.Bottom
						wordsAlignment = CellVerticalAlignment.Bottom
						Exit Select
				Case TextAlignmentType.Center
						wordsAlignment = CellVerticalAlignment.Center
						Exit Select
				Case TextAlignmentType.Top
						wordsAlignment = CellVerticalAlignment.Top
						Exit Select
				Case Else
						wordsAlignment = CellVerticalAlignment.Top
						Exit Select
			End Select

			Return wordsAlignment
		End Function

		''' <summary>
		''' Conver Excel HorizontalAlignment to Word ParagraphAlignment
		''' </summary>
		''' <param name="alignmentType">Excel AlignmentType</param>
		''' <returns>Word ParagraphAlignment</returns>
		Private Function ConvertHorizontalAlignment(ByVal alignmentType As TextAlignmentType) As ParagraphAlignment
			Dim wordsAlignment As ParagraphAlignment = ParagraphAlignment.Left

			Select Case alignmentType
				Case TextAlignmentType.Center
						wordsAlignment = ParagraphAlignment.Center
						Exit Select
				Case TextAlignmentType.Distributed
						wordsAlignment = ParagraphAlignment.Distributed
						Exit Select
				Case TextAlignmentType.Justify
						wordsAlignment = ParagraphAlignment.Justify
						Exit Select
				Case TextAlignmentType.Left
						wordsAlignment = ParagraphAlignment.Left
						Exit Select
				Case TextAlignmentType.Right
						wordsAlignment = ParagraphAlignment.Right
						Exit Select
				Case Else
						wordsAlignment = ParagraphAlignment.Left
						Exit Select
			End Select

			Return wordsAlignment
		End Function

		''' <summary>
		''' Convert Excel HorizontalCellMerge to Word HorizontalCellMerge
		''' </summary>
		''' <param name="excelCell">Input Excel cell</param>
		''' <returns>CellMerge type</returns>
		Private Function ConvertHorizontalCellMerge(ByVal excelCell As Aspose.Cells.Cell) As CellMerge
			'By default cells are not merged
			Dim wordsCellMerge As CellMerge = CellMerge.None
			'Get merged region
			Dim mergedRange As Aspose.Cells.Range = excelCell.GetMergedRange()
			If mergedRange Is Nothing Then
				'Cells are not merged
				wordsCellMerge = CellMerge.None
			Else
				If excelCell.Column = mergedRange.FirstColumn AndAlso mergedRange.ColumnCount > 1 Then
					'Cell is merged with next
					wordsCellMerge = CellMerge.First
				ElseIf mergedRange.ColumnCount > 1 Then
					'Cell is merged with previouse
					wordsCellMerge = CellMerge.Previous
				Else
					'Cell is not merged
					wordsCellMerge = CellMerge.None
				End If
			End If
			Return wordsCellMerge
		End Function

		''' <summary>
		''' Convert Excel VerticalCellMerge to Word VerticalCellMerge
		''' </summary>
		''' <param name="excelCell">Input Excel cell</param>
		''' <returns>CellMerge type</returns>
		Private Function ConvertVerticalCellMerge(ByVal excelCell As Aspose.Cells.Cell) As CellMerge
			'By default cells are not merged
			Dim wordsCellMerge As CellMerge = CellMerge.None
			'Get merged region
			Dim mergedRange As Aspose.Cells.Range = excelCell.GetMergedRange()

			If mergedRange Is Nothing Then
				'Cells are not merged
				wordsCellMerge = CellMerge.None
			Else
				If (excelCell.Row.Equals(mergedRange.FirstRow)) AndAlso (mergedRange.RowCount > 1) Then
					'Cell is merged with next
					wordsCellMerge = CellMerge.First
				ElseIf mergedRange.RowCount > 1 Then
					'Cell is merged with previouse
					wordsCellMerge = CellMerge.Previous
				Else
					'Cell is not merged
					wordsCellMerge = CellMerge.None
				End If
			End If
			Return wordsCellMerge
		End Function

		''' <summary>
		''' Convert Excel MsoLineStyle to Word ShapeLineStyle
		''' </summary>
		''' <param name="lineStyle">Excel MsoLineStyle</param>
		''' <returns>Word ShapeLineStyle</returns>
		Private Function ConvertDrawingLineStyle(ByVal lineStyle As MsoLineStyle) As Aspose.Words.Drawing.ShapeLineStyle
			Dim wordsLineStyle As Aspose.Words.Drawing.ShapeLineStyle = Aspose.Words.Drawing.ShapeLineStyle.Single

			Select Case lineStyle
				Case MsoLineStyle.Single
						wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.Single
						Exit Select
				Case MsoLineStyle.ThickBetweenThin
						wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.Triple
						Exit Select
				Case MsoLineStyle.ThickThin
						wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.ThickThin
						Exit Select
				Case MsoLineStyle.ThinThick
						wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.ThinThick
						Exit Select
				Case MsoLineStyle.ThinThin
						wordsLineStyle = Aspose.Words.Drawing.ShapeLineStyle.Double
						Exit Select
			End Select

			Return wordsLineStyle
		End Function

		''' <summary>
		''' Convert Excel MsoLineDashStyle to Word DashStyle
		''' </summary>
		''' <param name="dashStyle">Excel MsoLineDashStyle</param>
		''' <returns>Word DashStyle</returns>
		Private Function ConvertDrawingDashStyle(ByVal dashStyle As MsoLineDashStyle) As Aspose.Words.Drawing.DashStyle
			Dim wordsDashStyle As Aspose.Words.Drawing.DashStyle = Aspose.Words.Drawing.DashStyle.Solid
			Select Case dashStyle
				Case MsoLineDashStyle.Dash
						wordsDashStyle = Aspose.Words.Drawing.DashStyle.Dash
						Exit Select
				Case MsoLineDashStyle.DashDot
						wordsDashStyle = Aspose.Words.Drawing.DashStyle.DashDot
						Exit Select
				Case MsoLineDashStyle.DashDotDot
						wordsDashStyle = Aspose.Words.Drawing.DashStyle.LongDashDotDot
						Exit Select
				Case MsoLineDashStyle.DashLongDash
						wordsDashStyle = Aspose.Words.Drawing.DashStyle.LongDash
						Exit Select
				Case MsoLineDashStyle.DashLongDashDot
						wordsDashStyle = Aspose.Words.Drawing.DashStyle.LongDashDot
						Exit Select
				Case MsoLineDashStyle.RoundDot
						wordsDashStyle = Aspose.Words.Drawing.DashStyle.Dot
						Exit Select
				Case MsoLineDashStyle.Solid
						wordsDashStyle = Aspose.Words.Drawing.DashStyle.Solid
						Exit Select
				Case MsoLineDashStyle.SquareDot
						wordsDashStyle = Aspose.Words.Drawing.DashStyle.ShortDot
						Exit Select
			End Select
			Return wordsDashStyle
		End Function

		''' <summary>
		''' Convert Excel TextOrientationType to Word LayoutFlow
		''' </summary>
		''' <param name="textOrientation">Excel TextOrientationType</param>
		''' <returns>Word LayoutFlow</returns>
		Private Function ConvertDrawingTextOrientationType(ByVal textOrientation As TextOrientationType) As Aspose.Words.Drawing.LayoutFlow
			Dim wordLayoutFlow As Aspose.Words.Drawing.LayoutFlow = Aspose.Words.Drawing.LayoutFlow.Horizontal
			Select Case textOrientation
				Case TextOrientationType.ClockWise
						wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.BottomToTop
						Exit Select
				Case TextOrientationType.CounterClockWise
						wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.TopToBottom
						Exit Select
				Case TextOrientationType.NoRotation
						wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.Horizontal
						Exit Select
				Case TextOrientationType.TopToBottom
						wordLayoutFlow = Aspose.Words.Drawing.LayoutFlow.TopToBottom
						Exit Select
			End Select
			Return wordLayoutFlow
		End Function

		''' <summary>
		''' Convert Excel MsoDrawingType to Word ShapeType
		''' </summary>
		''' <param name="excelShapeType">Excel MsoDrawingType</param>
		''' <returns>Word ShapeType</returns>
		Private Function ConvertDrawingShapetype(ByVal excelShapeType As MsoDrawingType) As Aspose.Words.Drawing.ShapeType
			Dim wordsShapeType As Aspose.Words.Drawing.ShapeType = Aspose.Words.Drawing.ShapeType.Line
			Select Case excelShapeType
				Case MsoDrawingType.Arc
						wordsShapeType = Aspose.Words.Drawing.ShapeType.Arc
						Exit Select
				Case MsoDrawingType.Line
						wordsShapeType = Aspose.Words.Drawing.ShapeType.Line
						Exit Select
				Case MsoDrawingType.Oval
						wordsShapeType = Aspose.Words.Drawing.ShapeType.Ellipse
						Exit Select
				Case MsoDrawingType.Rectangle
						wordsShapeType = Aspose.Words.Drawing.ShapeType.Rectangle
						Exit Select
			End Select
			Return wordsShapeType
		End Function

		''' <summary>
		''' Convert excel BackgroundType to Word TextureIndex
		''' </summary>
		''' <param name="excelTextureType">excel BackgroundType</param>
		''' <returns>Word TextureIndex</returns>
		Private Function ConvertBackgroundTexture(ByVal excelTextureType As Aspose.Cells.BackgroundType) As Aspose.Words.TextureIndex
			Dim wordsTextureIndex As Aspose.Words.TextureIndex = Aspose.Words.TextureIndex.TextureNone
			Select Case excelTextureType
				Case BackgroundType.DiagonalCrosshatch
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalCross
						Exit Select
				Case BackgroundType.DiagonalStripe
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalDown
						Exit Select
				Case BackgroundType.Gray12
						wordsTextureIndex = Aspose.Words.TextureIndex.Texture12Pt5Percent
						Exit Select
				Case BackgroundType.Gray25
						wordsTextureIndex = Aspose.Words.TextureIndex.Texture25Percent
						Exit Select
				Case BackgroundType.Gray50
						wordsTextureIndex = Aspose.Words.TextureIndex.Texture50Percent
						Exit Select
				Case BackgroundType.Gray6
						wordsTextureIndex = Aspose.Words.TextureIndex.Texture10Percent
						Exit Select
				Case BackgroundType.Gray75
						wordsTextureIndex = Aspose.Words.TextureIndex.Texture75Percent
						Exit Select
				Case BackgroundType.HorizontalStripe
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureHorizontal
						Exit Select
				Case BackgroundType.None
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureNone
						Exit Select
				Case BackgroundType.ReverseDiagonalStripe
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalUp
						Exit Select
				Case BackgroundType.Solid
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureSolid
						Exit Select
				Case BackgroundType.ThickDiagonalCrosshatch
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureDarkDiagonalCross
						Exit Select
				Case BackgroundType.ThinDiagonalCrosshatch
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalCross
						Exit Select
				Case BackgroundType.ThinDiagonalStripe
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalCross
						Exit Select
				Case BackgroundType.ThinHorizontalCrosshatch
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureHorizontal
						Exit Select
				Case BackgroundType.ThinHorizontalStripe
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureHorizontal
						Exit Select
				Case BackgroundType.ThinReverseDiagonalStripe
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureDiagonalUp
						Exit Select
				Case BackgroundType.ThinVerticalStripe
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureVertical
						Exit Select
				Case BackgroundType.VerticalStripe
						wordsTextureIndex = Aspose.Words.TextureIndex.TextureVertical
						Exit Select

			End Select
			Return wordsTextureIndex
		End Function

		''' <summary>
		''' Get text with formating from Excel cell
		''' </summary>
		''' <param name="cell">Input Excel cell</param>
		''' <param name="doc">Parent document</param>
		''' <returns>Run that contains text with formating</returns>
		Private Function GetTextFromCell(ByVal cell As Aspose.Cells.Cell, ByVal doc As Document) As ArrayList
			'Create new array list 
			'We will store runs in this list
			Dim wordRuns As New ArrayList()

			'Get Chracters objects from the cell
			Dim charactersList() As FontSetting = cell.GetCharacters()
			If charactersList Is Nothing Then
				charactersList = New FontSetting(){}
			End If

			'Get sring value from the cell
			Dim cellValue As String = cell.StringValue

			'If there is some formatig in the cell 
			'charactersList will contains one or more Characters objects
			'And we should create collection of Runs that will represent this formating
			If charactersList.Length > 0 Then
				Dim startIndex As Integer = (CType(charactersList(0), FontSetting)).StartIndex
				If startIndex > 0 Then
					'Create first run
					Dim firstRun As New Run(doc)
					'Set text of first run. That will be substring that starts from 0
					'And ens and the start position of first Characters object

					'Sometimes Aspose.Cell returns incorrect index so we should check it
					If startIndex < cellValue.Length Then
						firstRun.Text = cellValue.Substring(0, startIndex)
					Else
						firstRun.Text = cellValue.Substring(0)
					End If
					'Formation of first run will be the same as formating of whole cell
					ImportFont(firstRun.Font, cell.GetDisplayStyle().Font)
					wordRuns.Add(firstRun)
				End If

				'Loop through all Character objects
				For Each chars As FontSetting In charactersList
					Dim run As New Run(doc)
					'We should check index to avoid errors
					If chars.StartIndex < cellValue.Length Then
						If cellValue.Length > (chars.StartIndex + chars.Length) AndAlso chars.Length > 0 Then
							run.Text = cellValue.Substring(chars.StartIndex, chars.Length)
						Else
							run.Text = cellValue.Substring(chars.StartIndex)
						End If
					End If
					'Convert Excel Font to Words Font
					ImportFont(run.Font, chars.Font)
					wordRuns.Add(run)
				Next chars
			'Othervice there will be only one run 
			Else
				Dim run As New Run(doc)
				run.Text = cellValue
				'Convert Excel Font to Words Font
				ImportFont(run.Font, cell.GetDisplayStyle().Font)
				wordRuns.Add(run)
			End If

			Return wordRuns
		End Function

		''' <summary>
		''' Convert Excel Font to Word Font
		''' </summary>
		''' <param name="wordsFont">Word Font</param>
		''' <param name="excelFont">Excel Font</param>
		Private Sub ImportFont(ByVal wordsFont As Aspose.Words.Font, ByVal excelFont As Aspose.Cells.Font)
			wordsFont.Name = excelFont.Name
			wordsFont.Bold = excelFont.IsBold
			wordsFont.Color = excelFont.Color
			wordsFont.Italic = excelFont.IsItalic
			wordsFont.StrikeThrough = excelFont.IsStrikeout
			wordsFont.Subscript = excelFont.IsSubscript
			wordsFont.Superscript = excelFont.IsSuperscript
			wordsFont.Size = excelFont.Size
			wordsFont.Underline = ConvertUnderline(excelFont.Underline)
		End Sub

		''' <summary>
		''' Convert Excel PaperSize to Word PaperSize
		''' </summary>
		''' <param name="excelPaperSize">Excel PaperSize</param>
		''' <returns>Word Paper size</returns>
		Private Function ConvertPaperSize(ByVal excelPaperSize As PaperSizeType) As PaperSize
			Dim paperSize As PaperSize = PaperSize.A4

			Select Case excelPaperSize
				Case PaperSizeType.PaperA4
						paperSize = PaperSize.A4
						Exit Select
				Case PaperSizeType.PaperA3
						paperSize = PaperSize.A3
						Exit Select
				Case PaperSizeType.PaperA5
						paperSize = PaperSize.A5
						Exit Select
				Case PaperSizeType.PaperB4
						paperSize = PaperSize.B4
						Exit Select
				Case PaperSizeType.PaperB5
						paperSize = PaperSize.B5
						Exit Select
				Case PaperSizeType.Paper10x14
						paperSize = PaperSize.Paper10x14
						Exit Select
				Case PaperSizeType.Paper11x17
						paperSize = PaperSize.Paper11x17
						Exit Select
				Case PaperSizeType.PaperEnvelopeDL
						paperSize = PaperSize.EnvelopeDL
						Exit Select
				Case PaperSizeType.PaperExecutive
						paperSize = PaperSize.Executive
						Exit Select
				Case PaperSizeType.PaperFolio
						paperSize = PaperSize.Folio
						Exit Select
				Case PaperSizeType.PaperLedger
						paperSize = PaperSize.Ledger
						Exit Select
				Case PaperSizeType.PaperLegal
						paperSize = PaperSize.Legal
						Exit Select
				Case PaperSizeType.PaperLetter
						paperSize = PaperSize.Letter
						Exit Select
				Case PaperSizeType.PaperQuarto
						paperSize = PaperSize.Quarto
						Exit Select
				Case PaperSizeType.PaperStatement
						paperSize = PaperSize.Statement
						Exit Select
				Case PaperSizeType.PaperTabloid
						paperSize = PaperSize.Tabloid
						Exit Select
				Case Else
						paperSize = PaperSize.Letter
						Exit Select
			End Select

			Return paperSize
		End Function

		''' <summary>
		''' Convert Excel PageOrientaton to Word PageOrientaton
		''' </summary>
		''' <param name="excelPageOrientation">Excel PageOrientation (Portrait or Landscape)</param>
		''' <returns>Portrait or Landscape, by default returns Portrait</returns>
		Private Function ConvertPageOrientation(ByVal excelPageOrientation As PageOrientationType) As Orientation
			Dim pageOrientation As Orientation = Orientation.Portrait

			Select Case excelPageOrientation
				Case PageOrientationType.Portrait
						pageOrientation = Orientation.Portrait
						Exit Select
				Case PageOrientationType.Landscape
						pageOrientation = Orientation.Landscape
						Exit Select
				Case Else
						pageOrientation = Orientation.Portrait
						Exit Select
			End Select

			Return pageOrientation
		End Function

		''' <summary>
		''' Calculate additional offset if excel cell is merged horizontally
		''' </summary>
		''' <param name="excelCell">Excel cell</param>
		''' <returns></returns>
		Private Function GetAdditionalHorizontalOffset(ByVal excelCell As Aspose.Cells.Cell) As Double
			Dim leftOffset As Double = 0
			'Get merged region of excel Cell
			Dim mergedRange As Aspose.Cells.Range = excelCell.GetMergedRange()
			If mergedRange IsNot Nothing Then
				If excelCell.Column <> mergedRange.FirstColumn AndAlso mergedRange.ColumnCount > 1 Then
					'Cell is merged with previouse
					For columnIndex As Integer = mergedRange.FirstColumn To excelCell.Column - 1
						leftOffset += ConvertUtil.PixelToPoint(mergedRange.Worksheet.Cells.GetColumnWidthPixel(columnIndex))
					Next columnIndex
				End If
			End If
			Return leftOffset
		End Function

		''' <summary>
		''' Calculate additional offset if excel cell is merged vertically
		''' </summary>
		''' <param name="excelCell">Excel cell</param>
		''' <returns></returns>
		Private Function GetAdditionalVerticalOffset(ByVal excelCell As Aspose.Cells.Cell) As Double
			Dim topOffset As Double = 0
			'Get merged region of excel Cell
			Dim mergedRange As Aspose.Cells.Range = excelCell.GetMergedRange()
			If mergedRange IsNot Nothing Then
				If ((Not excelCell.Row.Equals(mergedRange.FirstRow))) AndAlso (mergedRange.RowCount > 1) Then
					'Cell is merged with previouse
					For rowIndex As Integer = mergedRange.FirstRow To excelCell.Row - 1
						topOffset += ConvertUtil.PixelToPoint(mergedRange.Worksheet.Cells.GetRowHeightPixel(rowIndex))
					Next rowIndex
				End If
			End If
			Return topOffset
		End Function

		''' <summary>
		''' Convert Excel Picture to Word Shape
		''' </summary>
		''' <param name="excelPicture">Excel Picture</param>
		''' <param name="doc">Parent document</param>
		''' <returns>Word Shape</returns>
		Private Function ConvertPictureToShape(ByVal excelPicture As Aspose.Cells.Drawing.Picture, ByVal doc As DocumentBase) As Aspose.Words.Drawing.Shape
			'Create new Shape
			Dim wordsShape As New Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.Image)
			'Set image
			wordsShape.ImageData.ImageBytes = excelPicture.Data
			'Import Picture properties inhereted from Shape
			ImportShapeProperties(wordsShape, CType(excelPicture, Aspose.Cells.Drawing.Shape))
			Return wordsShape
		End Function

		''' <summary>
		''' Convert Excel Chart to Word Shape
		''' </summary>
		''' <param name="excelChart">Excel Chart</param>
		''' <param name="doc">Parent document</param>
		''' <returns>Word Shape</returns>
		Private Function ConvertCartToShape(ByVal excelChart As Aspose.Cells.Charts.Chart, ByVal doc As DocumentBase) As Aspose.Words.Drawing.Shape
			'Create a new Shape
			Dim wordsShape As New Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.Image)
			'Convert Chart to Bitmap. Now only supports to convert 2D chart to image. If the chart is 3D chart,return null. 
			Dim chartPicture As Bitmap = excelChart.ToImage()
			If chartPicture IsNot Nothing Then
				wordsShape.ImageData.SetImage(chartPicture)
				'Import Chart properties inhereted from Shape
				ImportShapeProperties(wordsShape, CType(excelChart.ChartObject, Shape))
				Return wordsShape
			Else
				Return Nothing
			End If
		End Function

		''' <summary>
		''' Insert CheckBox into a Word cell
		''' </summary>
		''' <param name="excelCheckbox">Excel CheckBox</param>
		''' <param name="parentCell">Parent Word cell</param>
		Private Sub InsertCheckBox(ByVal excelCheckbox As Aspose.Cells.Drawing.CheckBox, ByVal parentCell As Aspose.Words.Tables.Cell)
			'Create new temporary document
			Dim doc As New Document()
			'Create instance of DocumentBuilder
			Dim builder As New DocumentBuilder(doc)
			'Calculate size of CheckBox
			Dim size As Integer = CInt(Fix(ConvertUtil.PixelToPoint(excelCheckbox.Height)))

			Select Case excelCheckbox.CheckedValue
				Case CheckValueType.Checked
						builder.InsertCheckBox(excelCheckbox.Name, True, size)
						Exit Select
				Case CheckValueType.UnChecked
						builder.InsertCheckBox(excelCheckbox.Name, False, size)
						Exit Select
				Case Else
						builder.InsertCheckBox(excelCheckbox.Name, False, size)
						Exit Select
			End Select
			'Write text of Excel CheckBox
			builder.Write(excelCheckbox.Text)

			'Import all content of temporary document into a destination cell
			For Each node As Node In builder.CurrentParagraph.ChildNodes
				parentCell.LastParagraph.AppendChild(parentCell.Document.ImportNode(node, True))
			Next node
		End Sub

		''' <summary>
		''' Convert Excel TextBox to Word TextBox
		''' </summary>
		''' <param name="excelTextBox">Excel TextBox</param>
		''' <param name="doc">Parent document</param>
		''' <returns>Word Shape</returns>
		Private Function ConvertTextBoxToShape(ByVal excelTextBox As Aspose.Cells.Drawing.TextBox, ByVal doc As DocumentBase) As Aspose.Words.Drawing.Shape
			'Create a new TextBox
			Dim wordsShape As New Aspose.Words.Drawing.Shape(doc, Aspose.Words.Drawing.ShapeType.TextBox)
			'Import TextBox properties inhereted from Shape
			ImportShapeProperties(wordsShape, CType(excelTextBox, Shape))
			'Import TextBox properties
			wordsShape.TextBox.LayoutFlow = ConvertDrawingTextOrientationType(excelTextBox.TextOrientationType)
			'Import text
			Dim run As New Run(doc)
			If (Not String.IsNullOrEmpty(excelTextBox.Text)) Then
				run.Text = excelTextBox.Text
			Else
				run.Text = String.Empty
			End If
			'Import text formating
			ImportFont(run.Font, excelTextBox.Font)
			'Create paragraph
			Dim paragraph As New Paragraph(doc)
			'Import horizontal alignment
			paragraph.ParagraphFormat.Alignment = ConvertHorizontalAlignment(excelTextBox.TextHorizontalAlignment)
			'Insert text into the paragraph
			paragraph.AppendChild(run)
			'insert Pragraph into textbox
			wordsShape.AppendChild(paragraph)
			Return wordsShape
		End Function

		''' <summary>
		''' Convert Excel Shape to Word Shape
		''' </summary>
		''' <param name="excelShape">Excel Shape</param>
		''' <param name="doc">Parent document</param>
		''' <returns>Word Shape</returns>
		Private Function ConvertShapeToShape(ByVal excelShape As Aspose.Cells.Drawing.Shape, ByVal doc As DocumentBase) As Aspose.Words.Drawing.Shape
			'Create words Shape
			Dim wordsShape As New Aspose.Words.Drawing.Shape(doc, ConvertDrawingShapetype(excelShape.MsoDrawingType))
			'Import properties
			ImportShapeProperties(wordsShape, excelShape)

			wordsShape.Stroked = True
			wordsShape.Filled = True

			Return wordsShape
		End Function

		''' <summary>
		''' Import properties of Excel shape
		''' </summary>
		''' <param name="wordsShape">Word Shape</param>
		''' <param name="excelShape">Excel Shape</param>
		Private Sub ImportShapeProperties(ByVal wordsShape As Aspose.Words.Drawing.Shape, ByVal excelShape As Aspose.Cells.Drawing.Shape)
			'Import size of TextBox
			wordsShape.Height = ConvertUtil.PixelToPoint(excelShape.Height) '1pt=1px*0.75
			wordsShape.Width = ConvertUtil.PixelToPoint(excelShape.Width)
			'Import horizontal offset
			wordsShape.Left = ConvertUtil.PixelToPoint(excelShape.Left)
			'Import vertical offset
			wordsShape.Top = ConvertUtil.PixelToPoint(excelShape.Top)

			'Import Filling
			If excelShape.FillFormat.IsVisible Then
				wordsShape.Filled = True
				wordsShape.Fill.Color = excelShape.FillFormat.ForeColor
			Else
				wordsShape.Filled = False
			End If
			'Import LineFormat (borders)
			If excelShape.LineFormat.IsVisible Then
				wordsShape.Stroked = True
				'Set LineStyle
				wordsShape.Stroke.LineStyle = ConvertDrawingLineStyle(excelShape.LineFormat.Style)
				'Set DashStyle
				wordsShape.Stroke.DashStyle = ConvertDrawingDashStyle(excelShape.LineFormat.DashStyle)
				'Set Weight
				wordsShape.Stroke.Weight = excelShape.LineFormat.Weight
				'Set collors
				wordsShape.Stroke.Color = If(excelShape.LineFormat.ForeColor.IsEmpty, Color.Black, excelShape.LineFormat.ForeColor)
				wordsShape.Stroke.Color2 = excelShape.LineFormat.BackColor
			Else
				wordsShape.Stroked = False
			End If
			'Import link
			If excelShape.Hyperlink IsNot Nothing Then
				wordsShape.HRef = excelShape.Hyperlink.Address
			End If
			'Import rotation
			wordsShape.Rotation = excelShape.RotationAngle
		End Sub


		#Region "Private variables"

		'Create HashTables. We will store in these tables objects like Pictures, Charts, etc.
		Private mShapesCollection As New Hashtable()
		Private mPicturesCollection As New Hashtable()
		Private mChartsCollection As New Hashtable()
		Private mTextBoxesCollection As New Hashtable()
		Private mCheckBoxesCollection As New Hashtable()

		#End Region
	End Class
End Namespace
