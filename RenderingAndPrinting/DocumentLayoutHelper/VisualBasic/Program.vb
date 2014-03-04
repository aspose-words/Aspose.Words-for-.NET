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
Imports Aspose.Words.Tables

Namespace DocumentLayoutHelperExample
	Public Class Program
		Public Shared Sub Main()
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim doc As New Document(dataDir & "TestFile.docx")

			' This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for 
			' the LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.

			' Create a new RenderedDocument class from a Document object.
			Dim layoutDoc As New RenderedDocument(doc)

			' The following examples demonstrate how to use the wrapper API. 
			' This snippet returns the third line of the first page and prints the line of text to the console.
			Dim line As RenderedLine = layoutDoc.Pages(0).Columns(0).Lines(2)
			Console.WriteLine("Line: " & line.Text)

			' With a rendered line the original paragraph in the document object model can be returned.
			Dim para As Paragraph = line.Paragraph
			Console.WriteLine("Paragraph text: " & para.Range.Text)

			' Retrieve all the text that appears of the first page in plain text format (including headers and footers).
			Dim pageText As String = layoutDoc.Pages(0).Text
			Console.WriteLine()

			' Loop through each page in the document and print how many lines appear on each page.
			For Each page As RenderedPage In layoutDoc.Pages
				Dim lines As LayoutCollection(Of LayoutEntity) = page.GetChildEntities(LayoutEntityType.Line, True)
				Console.WriteLine("Page {0} has {1} lines.", page.PageIndex, lines.Count)
			Next page

			' This method provides a reverse lookup of layout entities for any given node (with the exception of runs and nodes in the
			' header and footer).
			Console.WriteLine()
			Console.WriteLine("The lines of the second paragraph:")
			For Each paragraphLine As RenderedLine In layoutDoc.GetLayoutEntitiesOfNode(doc.FirstSection.Body.Paragraphs(1))
				Console.WriteLine(String.Format("""{0}""", paragraphLine.Text.Trim()))
				Console.WriteLine(paragraphLine.Rectangle.ToString())
				Console.WriteLine()
			Next paragraphLine
		End Sub
	End Class
End Namespace