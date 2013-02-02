'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Net
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Shapes
Imports System.Windows.Navigation
Imports System.IO
Imports System.Windows.Browser

Namespace FirstFloor.Documents.Aspose.Views
	Partial Public Class Demo
		Inherits Page
		Public Sub New()
			InitializeComponent()

			AddHandler explorer.DocumentSelected, AddressOf explorer_DocumentSelected
		End Sub

		Private Sub explorer_DocumentSelected(ByVal sender As Object, ByVal e As DocumentSelectedEventArgs)
			Me.PageScrollViewer.Visibility = Visibility.Collapsed
			Me.viewer.Visibility = Visibility.Visible

			Me.viewer.LoadDocument(e.Document.Name, e.Document.XpsLocation, e.Document.OriginalLocation)
		End Sub

		' Executes when the user navigates to this page.
		Protected Overrides Sub OnNavigatedTo(ByVal e As NavigationEventArgs)
		End Sub

		Private Sub Button_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			Dim dlg = New OpenFileDialog()
			dlg.Multiselect = False
			dlg.Filter = "Word documents (*.doc,*.docx)|*.doc;*.docx|WordML documents (*.xml)|*.xml|OpenDocument Text (*.odt)|*.odt|HTML pages (*.htm,*.html)|*.htm;*.html|RTF documents (*.rtf)|*.rtf|All files (*.*)|*.*"

			If True = dlg.ShowDialog() Then
				If dlg.File.Length > 2 << 18 Then
					MessageBox.Show("The selected document is too large. This demo limits the file size to 512KB." & Constants.vbLf + Constants.vbLf & "Please select a smaller document.", "Document too large", MessageBoxButton.OK)
					Return
				End If
				Me.PageScrollViewer.Visibility = Visibility.Collapsed
				Me.viewer.Visibility = Visibility.Visible
				Me.viewer.ClearDocument()

				Me.viewer.LoadLocalDocument(dlg.File)
			End If
		End Sub

		Private Sub viewer_Close(ByVal sender As Object, ByVal e As EventArgs)
			Me.PageScrollViewer.Visibility = Visibility.Visible
			Me.viewer.Visibility = Visibility.Collapsed
		End Sub
	End Class
End Namespace
