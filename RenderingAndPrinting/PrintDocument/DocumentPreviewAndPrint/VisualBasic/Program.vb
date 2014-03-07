'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms

Imports Aspose.Words
Imports Aspose.Words.Rendering

Namespace DocumentPreviewAndPrintExample
	'ExStart
	'ExId:DocumentPreviewAndPrint_ActivePrintPreviewDialog_OnShown
	'ExSummary:Brings the Print Preview dialog to the front.
	''' <summary>
	''' A subclass of the .NET Print Preview dialog. This extension is used only to work around the .NET PrintPreviewDialog
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
	'ExEnd

	''' <summary>
	''' This project is set to target the x86 platform because the .NET print dialog does not 
	''' seem to show when calling from a 64-bit application.
	''' </summary>
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Open the document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			'ExStart
			'ExId:DocumentPreviewAndPrint_PrintDialog_Creation
			'ExSummary:Creates the print dialog.
			Dim printDlg As New PrintDialog()
			' Initialize the print dialog with the number of pages in the document.
			printDlg.AllowSomePages = True
			printDlg.PrinterSettings.MinimumPage = 1
			printDlg.PrinterSettings.MaximumPage = doc.PageCount
			printDlg.PrinterSettings.FromPage = 1
			printDlg.PrinterSettings.ToPage = doc.PageCount
			'ExEnd

			'ExStart
			'ExId:DocumentPreviewAndPrint_PrintDialog_Check_Result
			'ExSummary:Check if the user accepted the print settings and proceed to preview the document.
			If (Not printDlg.ShowDialog().Equals(DialogResult.OK)) Then
				Return
			End If
			'ExEnd

			'ExStart
			'ExId:DocumentPreviewAndPrint_AsposeWordsPrintDocument_Creation
			'ExSummary:Creates a special Aspose.Words implementation of the .NET PrintDocument class.
			' Pass the printer settings from the dialog to the print document.
			Dim awPrintDoc As New AsposeWordsPrintDocument(doc)
			awPrintDoc.PrinterSettings = printDlg.PrinterSettings
			'ExEnd

			'ExStart
			'ExId:DocumentPreviewAndPrint_ActivePrintPreviewDialog_Creation
			'ExSummary:Creates an overridden version of the .NET Print Preview dialog to preview the document.
			Dim previewDlg As New ActivePrintPreviewDialog()

			' Pass the Aspose.Words print document to the Print Preview dialog.
			previewDlg.Document = awPrintDoc
			' Specify additional parameters of the Print Preview dialog.
			previewDlg.ShowInTaskbar = True
			previewDlg.MinimizeBox = True
			previewDlg.PrintPreviewControl.Zoom = 1
			previewDlg.Document.DocumentName = "TestName.doc"
			previewDlg.WindowState = FormWindowState.Maximized
			' Show the appropriately configured Print Preview dialog.
			previewDlg.ShowDialog()
			'ExEnd
		End Sub
	End Class
End Namespace