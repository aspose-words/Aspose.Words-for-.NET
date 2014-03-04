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
Imports System.Windows.Forms
Imports Aspose.Cells
Imports Aspose.Words

Namespace Excel2WordExample
	Partial Public Class MainForm
		Inherits Form
		Public Sub New()
			InitializeComponent()
			Dim wordsLicenseFile As String = Path.Combine(Application.StartupPath, "Aspose.Words.lic")
			If File.Exists(wordsLicenseFile) Then
				'This shows how to license Aspose.Words.
				'If you don't specify a license, Aspose.Words works in evaluation mode.
				Dim license As New Aspose.Words.License()
				license.SetLicense(wordsLicenseFile)
			End If

			Dim cellsLicenseFile As String = Path.Combine(Application.StartupPath, "Aspose.Cells.lic")
			If File.Exists(cellsLicenseFile) Then
				'This shows how to license Aspose.Cells.
				'If you don't specify a license, Aspose.Cells works in evaluation mode.
				Dim license As New Aspose.Cells.License()
				license.SetLicense(cellsLicenseFile)
			End If
		End Sub

		Private Sub buttonConvert_Click(ByVal sender As Object, ByVal e As EventArgs) Handles buttonConvert.Click
			Try
				'Show the open dialog
				If (Not openFileDialog.ShowDialog().Equals(System.Windows.Forms.DialogResult.OK)) Then
					Return
				End If

				'Show the save dialog to select the destination file name and then run the demo.
				saveFileDialog.FileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName) & " Out"
				If (Not saveFileDialog.ShowDialog().Equals(System.Windows.Forms.DialogResult.OK)) Then
					Return
				End If

				RunConvert(openFileDialog.FileName, saveFileDialog.FileName)

				MessageBox.Show("Done!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information)
			Catch ex As Exception
				MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
			End Try
		End Sub

		Private Shared Sub RunConvert(ByVal srcFileName As String, ByVal dstFileName As String)
			'Open Excel Workbook using Aspose.Cells.
			Dim workbook As New Workbook(srcFileName)

			'Convert workbook to Word document
			Dim converter As New ConverterXls2Doc()
			Dim doc As Document = converter.Convert(workbook)

			' Save using Aspose.Words. 
			doc.Save(dstFileName)
		End Sub
	End Class
End Namespace