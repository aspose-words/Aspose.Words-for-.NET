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
Imports Aspose.Words
#If CellsInstalled Then
Imports Aspose.Cells
#End If

Namespace Excel2WordExample
	Partial Public Class MainForm
		Inherits Form
		Public Sub New()
			InitializeComponent()
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
			#If CellsInstalled Then
			'Open Excel Workbook using Aspose.Cells.
			Dim workbook As New Workbook(srcFileName)

			'Convert workbook to Word document
			Dim converter As New ConverterXls2Doc()
			Dim doc As Document = converter.Convert(workbook)

			' Save using Aspose.Words. 
			doc.Save(dstFileName)
#Else
			Throw New InvalidOperationException("This example requires the use of Aspose.Cells." & "Make sure Aspose.Cells.dll is present in the bin" & Constants.vbLf & "et2.0 folder.")
#End If
		End Sub

		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		<STAThread> _
		Public Shared Sub Main()
			Application.EnableVisualStyles()
			Application.SetCompatibleTextRenderingDefault(False)
			Application.Run(New MainForm())
		End Sub
	End Class
End Namespace