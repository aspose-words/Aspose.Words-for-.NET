'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Drawing.Text
Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms

Imports Aspose.Words

Namespace MultiplePagesOnSheetExample
	''' <summary>
	''' A subclass of the .NET Print Preview dialog. This extension only is used only to work around the .NET PrintPreviewDialog
	''' class not appearing in front of other windows by default.
	''' </summary>
	Friend Class ActivePrintPreviewDialog
		Inherits PrintPreviewDialog
		''' <summary>
		''' Brings the Print Preview dialog on top when it is initially displayed.
		''' </summary>
		Protected Overrides Sub OnShown(ByVal e As EventArgs)
			Activate()
			MyBase.OnShown(e)
		End Sub
	End Class

	''' <summary>
	''' This project is set to target the x86 platform because the .NET print dialog does not 
	''' seem to show when calling from a 64-bit application.
	''' </summary>
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'ExStart
			'ExId:MultiplePagesOnSheet_PrintAndPreview
			'ExSummary:The usage of the MultipagePrintDocument for Previewing and Printing.
			' Open the document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			Dim printDlg As New PrintDialog()
			' Initialize the Print Dialog with the number of pages in the document.
			printDlg.AllowSomePages = True
			printDlg.PrinterSettings.MinimumPage = 1
			printDlg.PrinterSettings.MaximumPage = doc.PageCount
			printDlg.PrinterSettings.FromPage = 1
			printDlg.PrinterSettings.ToPage = doc.PageCount

			' Check if user accepted the print settings and proceed to preview.
			If (Not printDlg.ShowDialog().Equals(DialogResult.OK)) Then
				Return
			End If

			' Pass the printer settings from the dialog to the print document.
			Dim awPrintDoc As New MultipagePrintDocument(doc, 4, True)
			awPrintDoc.PrinterSettings = printDlg.PrinterSettings

			' Create and configure the the ActivePrintPreviewDialog class.
			Dim previewDlg As New ActivePrintPreviewDialog()
			previewDlg.Document = awPrintDoc
			' Specify additional parameters of the Print Preview dialog.
			previewDlg.ShowInTaskbar = True
			previewDlg.MinimizeBox = True
			previewDlg.Document.DocumentName = "TestFile.doc"
			previewDlg.WindowState = FormWindowState.Maximized
			' Show appropriately configured Print Preview dialog.
			previewDlg.ShowDialog()
			'ExEnd
		End Sub
	End Class

	'ExStart
	'ExId:MultiplePagesOnSheet_PrintDocument
	'ExSummary:The custom PrintDocument class.
	Friend Class MultipagePrintDocument
		Inherits PrintDocument
	'ExEnd
		''' <summary>
		''' Initializes a new instance of this class.
		''' </summary>
		''' <param name="document">The document to print.</param>
		''' <param name="pagesPerSheet">The number of pages per one sheet.</param>
		''' <param name="printPageBorders">The flag that indicates if the printed page borders are rendered.</param>
		'ExStart
		'ExId:MultiplePagesOnSheet_Constructor
		'ExSummary:The constructor of the custom PrintDocument class.
		Public Sub New(ByVal document As Document, ByVal pagesPerSheet As Integer, ByVal printPageBorders As Boolean)
			If document Is Nothing Then
				Throw New ArgumentNullException("document")
			End If

			mDocument = document
			mPagesPerSheet = pagesPerSheet
			mPrintPageBorders = printPageBorders
		End Sub
		'ExEnd

		''' <summary>
		''' Called before the printing starts. Initializes the range of pages to be printed
		''' according to the user's selection.
		''' </summary>
		''' <param name="e">The event arguments.</param>
		'ExStart
		'ExId:MultiplePagesOnSheet_OnBeginPrint
		'ExSummary:The overridden method OnBeginPrint, which is called before the first page of the document prints.
		Protected Overrides Sub OnBeginPrint(ByVal e As PrintEventArgs)
			MyBase.OnBeginPrint(e)

			Select Case PrinterSettings.PrintRange
				Case PrintRange.AllPages
					mCurrentPage = 0
					mPageTo = mDocument.PageCount - 1
				Case PrintRange.SomePages
					mCurrentPage = PrinterSettings.FromPage - 1
					mPageTo = PrinterSettings.ToPage - 1
				Case Else
					Throw New InvalidOperationException("Unsupported print range.")
			End Select

			' Store the size of the paper, selected by user, taking into account the paper orientation.
			If PrinterSettings.DefaultPageSettings.Landscape Then
				mPaperSize = New Size(PrinterSettings.DefaultPageSettings.PaperSize.Height, PrinterSettings.DefaultPageSettings.PaperSize.Width)
			Else
				mPaperSize = New Size(PrinterSettings.DefaultPageSettings.PaperSize.Width, PrinterSettings.DefaultPageSettings.PaperSize.Height)
			End If

		End Sub
		'ExEnd

		''' <summary>
		''' Converts the pagesPerSheet number into the number of columns and rows.
		''' </summary>
		''' <param name="pagesPerSheet">The number of the pages to be printed on the one sheet of paper.</param>
		'ExStart
		'ExId:MultiplePagesOnSheet_GetThumbCount
		'ExSummary:Defines the number of columns and rows depending on the pagesPerSheet number and the page orientation.
		Private Function GetThumbCount(ByVal pagesPerSheet As Integer) As Size
			Dim size As Size
			' Define the number of the columns and rows on the sheet for the Landscape-oriented paper.
			Select Case pagesPerSheet
				Case 16
					size = New Size(4, 4)
				Case 9
					size = New Size(3, 3)
				Case 8
					size = New Size(4, 2)
				Case 6
					size = New Size(3, 2)
				Case 4
					size = New Size(2, 2)
				Case 2
					size = New Size(2, 1)
				Case Else 
					size = New Size(1, 1)
			End Select
			' Switch the width and height if the paper is in the Portrait orientation.
			If mPaperSize.Width < mPaperSize.Height Then
				Return New Size(size.Height, size.Width)
			End If
			Return size
		End Function
		'ExEnd

		''' <summary>
		''' Converts 1/100 inch into 1/72 inch (points).
		''' </summary>
		''' <param name="value">The 1/100 inch value to convert.</param>
		'ExStart
		'ExId:MultiplePagesOnSheet_HundredthsInchToPoint
		'ExSummary:Converts hundredths of inches to points.
		Private Shared Function HundredthsInchToPoint(ByVal value As Single) As Single
			Return CSng(ConvertUtil.InchToPoint(value / 100))
		End Function
		'ExEnd

		''' <summary>
		''' Called when each page is printed. This method actually renders the page to the graphics object.
		''' </summary>
		''' <param name="e">The event arguments.</param>
		'ExStart
		'ExId:MultiplePagesOnSheet_OnPrintPage
		'ExSummary:Generates the printed page from the specified number of the document pages.
		Protected Overrides Sub OnPrintPage(ByVal e As PrintPageEventArgs)
			MyBase.OnPrintPage(e)

			' Transfer to the point units.
			e.Graphics.PageUnit = GraphicsUnit.Point
			e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit

			' Get the number of the thumbnail placeholders across and down the paper sheet.
			Dim thumbCount As Size = GetThumbCount(mPagesPerSheet)

			' Calculate the size of each thumbnail placeholder in points.
			' Papersize in .NET is represented in hundreds of an inch. We need to convert this value to points first.
			Dim thumbSize As New SizeF(HundredthsInchToPoint(mPaperSize.Width) / thumbCount.Width, HundredthsInchToPoint(mPaperSize.Height) / thumbCount.Height)

			' Select the number of the last page to be printed on this sheet of paper.
			Dim pageTo As Integer = Math.Min(mCurrentPage + mPagesPerSheet - 1, mPageTo)

			' Loop through the selected pages from the stored current page to calculated last page.
			For pageIndex As Integer = mCurrentPage To pageTo
				' Calculate the column and row indices.
				Dim columnIdx As Integer
				Dim rowIdx As Integer = Math.DivRem(pageIndex - mCurrentPage, thumbCount.Width, columnIdx)

				' Define the thumbnail location in world coordinates (points in this case).
				Dim thumbLeft As Single = columnIdx * thumbSize.Width
				Dim thumbTop As Single = rowIdx * thumbSize.Height
				' Render the document page to the Graphics object using calculated coordinates and thumbnail placeholder size.
				' The useful return value is the scale at which the page was rendered.
				Dim scale As Single = mDocument.RenderToSize(pageIndex, e.Graphics, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height)

				' Draw the page borders (the page thumbnail could be smaller than the thumbnail placeholder size).
				If mPrintPageBorders Then
					' Get the real 100% size of the page in points.
					Dim pageSize As SizeF = mDocument.GetPageInfo(pageIndex).SizeInPoints
					' Draw the border around the scaled page using the known scale factor.
					e.Graphics.DrawRectangle(Pens.Black, thumbLeft, thumbTop, pageSize.Width * scale, pageSize.Height * scale)

					' Draws the border around the thumbnail placeholder.
					e.Graphics.DrawRectangle(Pens.Red, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height)
				End If
			Next pageIndex

			' Re-calculate next current page and continue with printing if such page resides within the print range.
			mCurrentPage = mCurrentPage + mPagesPerSheet
			e.HasMorePages = (mCurrentPage <= mPageTo)
		End Sub
		'ExEnd

		'ExStart
		'ExId:MultiplePagesOnSheet_Fields
		'ExSummary:The data and state fields of the custom PrintDocument class.
		Private ReadOnly mDocument As Document
		Private ReadOnly mPagesPerSheet As Integer
		Private ReadOnly mPrintPageBorders As Boolean
		Private mPaperSize As Size
		Private mCurrentPage As Integer
		Private mPageTo As Integer
		'ExEnd
	End Class
End Namespace