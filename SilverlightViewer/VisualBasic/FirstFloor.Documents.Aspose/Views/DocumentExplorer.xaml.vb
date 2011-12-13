'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
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
Imports System.Windows.Browser
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Shapes
Imports System.Xml.Linq

Namespace FirstFloor.Documents.Aspose.Views
	Partial Public Class DocumentExplorer
		Inherits UserControl
		Public Event DocumentSelected As EventHandler(Of DocumentSelectedEventArgs)

		Public Sub New()
			InitializeComponent()

			Dim documentsUri = New Uri(HtmlPage.Document.DocumentUri, "Documents/Documents.xml")

			Dim client = New WebClient()
			AddHandler client.OpenReadCompleted, Function(o, e)
				If (Not e.Cancelled) Then
					If e.Error IsNot Nothing Then
						ErrorWindow.ShowError(e.Error)
					Else
						Using e.Result
							Dim doc = XDocument.Load(e.Result, LoadOptions.None)
							Dim docs = From document In doc.Descendants("Document") _
							           Select New DocumentInfo() With {.Name = CStr(document.Attribute("Name")), .Description= CStr(document.Attribute("Description")), .OriginalLocation = New Uri(documentsUri, CStr(document.Attribute("Name"))), .XpsLocation = New Uri(documentsUri, CStr(document.Attribute("XpsLocation")))}

							Me.documents.ItemsSource = docs
						End Using
					End If
				End If

			client.OpenReadAsync(documentsUri)
		End Sub

		Private Sub HyperlinkButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			Dim button = CType(sender, HyperlinkButton)
			Dim document = CType(button.DataContext, DocumentInfo)

			RaiseEvent DocumentSelected(Me, New DocumentSelectedEventArgs(document))
		End Sub
	End Class
End Namespace
