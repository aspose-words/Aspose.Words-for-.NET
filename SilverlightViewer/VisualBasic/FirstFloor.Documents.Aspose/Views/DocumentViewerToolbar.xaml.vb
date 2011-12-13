'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Globalization
Imports System.Linq
Imports System.Windows
Imports System.Windows.Browser
Imports System.Windows.Controls
Imports System.Windows.Input

Imports FirstFloor.Documents.Controls

Namespace FirstFloor.Documents.Aspose.Views
	Partial Public Class DocumentViewerToolbar
		Inherits UserControl
		Private viewer_Renamed As FixedDocumentViewer
		Private originalUri_Renamed As Uri

		Public Sub New()
			InitializeComponent()

			Dim modes = ViewMode.GetDefaultItems()
			Me.viewMode.ItemsSource = modes
			Me.viewMode.SelectedItem = modes.First(Function(m) m.Scale = 1)
			Me.IsEnabled = False
		End Sub

		Public Property Viewer() As FixedDocumentViewer
			Get
				Return Me.viewer_Renamed
			End Get
			Set(ByVal value As FixedDocumentViewer)
				Me.viewer_Renamed = value
				Me.viewer_Renamed.ViewMode = CType(Me.viewMode.SelectedItem, ViewMode)
				Me.IsEnabled = Me.viewer_Renamed IsNot Nothing

				If Me.viewer_Renamed IsNot Nothing Then
					AddHandler Me.viewer_Renamed.PageNumberChanged, AddressOf viewer_PageNumberChanged
				End If
			End Set
		End Property

		Private Sub viewer_PageNumberChanged(ByVal sender As Object, ByVal e As PageNumberChangedEventArgs)
			SelectPage(e.PageNumber)
		End Sub

		Public Property OriginalUri() As Uri
			Get
				Return Me.originalUri_Renamed
			End Get
			Set(ByVal value As Uri)
				Me.originalUri_Renamed = value
				Me.download.Visibility = If(value IsNot Nothing, Visibility.Visible, Visibility.Collapsed)
			End Set
		End Property

		Public Sub Refresh()
			If Me.viewer_Renamed IsNot Nothing Then
				SelectPage(1)
				Me.IsEnabled = Me.viewer_Renamed.PageCount > 0
			End If
		End Sub

		Private Sub SelectPage(ByVal pageNumber As Integer)
			Dim pageCount = Me.viewer_Renamed.PageCount
			pageNumber = Math.Max(If(pageCount > 0, 1, 0), Math.Min(pageNumber, pageCount))

			Me.viewer_Renamed.PageNumber = pageNumber
			Me.page.Text = String.Format("{0}", pageNumber)
			Me.pages.Text = String.Format("/ {0}", pageCount)
			Me.next.IsEnabled = pageNumber < pageCount
			Me.prev.IsEnabled = pageNumber > 1
		End Sub

		Private Sub prev_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			SelectPage(Me.viewer_Renamed.PageNumber - 1)
		End Sub

		Private Sub next_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			SelectPage(Me.viewer_Renamed.PageNumber + 1)
		End Sub

		Private Sub page_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
			If e.Key = Key.Enter Then
				Try
					Dim pageNumber As Integer = Integer.Parse(page.Text, CultureInfo.InvariantCulture)
					SelectPage(pageNumber)
				Catch e1 As FormatException
					SelectPage(Me.viewer_Renamed.PageNumber)
				End Try
			End If
		End Sub

		Private Sub viewMode_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
			If Me.viewer_Renamed IsNot Nothing Then
				Me.viewer_Renamed.ViewMode = CType(viewMode.SelectedItem, ViewMode)

				' set focus back to viewer
				Me.viewer_Renamed.Focus()
			End If
		End Sub

		Private Sub screen_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			App.Current.Host.Content.IsFullScreen = Not App.Current.Host.Content.IsFullScreen
		End Sub

		Private Sub download_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			HtmlPage.Window.Navigate(Me.originalUri_Renamed, "_blank")
		End Sub
	End Class
End Namespace
