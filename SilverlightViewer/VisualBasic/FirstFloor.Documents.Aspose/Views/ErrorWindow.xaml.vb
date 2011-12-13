'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Windows
Imports System.Windows.Controls

Namespace FirstFloor.Documents.Aspose
	Partial Public Class ErrorWindow
		Inherits ChildWindow
		Public Sub New(ByVal e As Exception)
			InitializeComponent()
			If e IsNot Nothing Then
				ErrorTextBox.Text = e.Message & Environment.NewLine & Environment.NewLine & e.StackTrace
			End If
		End Sub

		Public Sub New(ByVal uri As Uri)
			InitializeComponent()
			If uri IsNot Nothing Then
				ErrorTextBox.Text = "Page not found: """ & uri.ToString() & """"
			End If
		End Sub

		Public Sub New(ByVal message As String, ByVal details As String)
			InitializeComponent()
			ErrorTextBox.Text = message & Environment.NewLine & Environment.NewLine & details
		End Sub

		Private Sub OKButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			Me.DialogResult = True
		End Sub

		Public Shared Sub ShowError(ByVal [error] As Exception)
			Dim wnd = New ErrorWindow([error])
			wnd.Show()
		End Sub
	End Class
End Namespace