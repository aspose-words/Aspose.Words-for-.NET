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

Public Class DocumentPreviewAndPrint
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        ' Open the document.
        Dim doc As New Document(dataDir & "TestFile.doc")

        Dim printDlg As New PrintDialog()
        ' Initialize the print dialog with the number of pages in the document.
        printDlg.AllowSomePages = True
        printDlg.PrinterSettings.MinimumPage = 1
        printDlg.PrinterSettings.MaximumPage = doc.PageCount
        printDlg.PrinterSettings.FromPage = 1
        printDlg.PrinterSettings.ToPage = doc.PageCount
        
        If (Not printDlg.ShowDialog().Equals(DialogResult.OK)) Then
            Return
        End If
        
        ' Pass the printer settings from the dialog to the print document.
        Dim awPrintDoc As New AsposeWordsPrintDocument(doc)
        awPrintDoc.PrinterSettings = printDlg.PrinterSettings
        
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
    End Sub
End Class
