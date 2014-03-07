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

Imports Aspose.Words
Imports Aspose.Words.Layout
Imports System.Collections
Imports Aspose.Words.Drawing

Namespace AddImageToEachPageExample
	Public Class Program
		Public Shared Sub Main()
			' This a document that we want to add an image and custom text for each page without using the header or footer.
			Dim doc As New Document(gDataDir & "TestFile.doc")

			' Create and attach collector before the document before page layout is built.
			Dim layoutCollector As New LayoutCollector(doc)

			' Images in a document are added to paragraphs, so to add an image to every page we need to find at any paragraph 
			' belonging to each page.
			Dim enumerator As IEnumerator = doc.SelectNodes("//Body/Paragraph").GetEnumerator()

			' Loop through each document page.
			For page As Integer = 1 To doc.PageCount
				Do While enumerator.MoveNext()
					' Check if the current paragraph belongs to the target page.
					Dim paragraph As Paragraph = CType(enumerator.Current, Paragraph)
					If layoutCollector.GetStartPageIndex(paragraph) = page Then
						AddImageToPage(paragraph, page)
						Exit Do
					End If
				Loop
			Next page

			doc.Save(gDataDir & "TestFile Out.docx")
		End Sub

		''' <summary>
		''' Adds an image to a page using the supplied paragraph.
		''' </summary>
		''' <param name="para">The paragraph to an an image to.</param>
		''' <param name="page">The page number the paragraph appears on.</param>
		Public Shared Sub AddImageToPage(ByVal para As Paragraph, ByVal page As Integer)
			Dim doc As Document = CType(para.Document, Document)

			Dim builder As New DocumentBuilder(doc)
			builder.MoveTo(para)

			' Add a logo to the top left of the page. The image is placed infront of all other text.
			Dim shape As Shape = builder.InsertImage(gDataDir & "Aspose Logo.png", RelativeHorizontalPosition.Page, 60, RelativeVerticalPosition.Page, 60, -1, -1, WrapType.None)

			' Add a textbox next to the image which contains some text consisting of the page number. 
			Dim textBox As New Shape(doc, ShapeType.TextBox)

			' We want a floating shape relative to the page.
			textBox.WrapType = WrapType.None
			textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page
			textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page

			' Set the textbox position.
			textBox.Height = 30
			textBox.Width = 200
			textBox.Left = 150
			textBox.Top = 80

			' Add the textbox and set text.
			textBox.AppendChild(New Paragraph(doc))
			builder.InsertNode(textBox)
			builder.MoveTo(textBox.FirstChild)
			builder.Writeln("This is a custom note for page " & page)
		End Sub

		Public Shared gDataDir As String = Path.GetFullPath("../../../Data/")
	End Class
End Namespace