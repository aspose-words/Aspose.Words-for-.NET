' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Tables

Namespace ApiExamples
	''' <summary>
	''' Functions for operations with document and content
	''' </summary>
	Friend NotInheritable Class DocumentHelper
		''' <summary>
		''' Create new document without run in the paragraph
		''' </summary>
		Private Sub New()
		End Sub
		Friend Shared Function CreateDocumentWithoutDummyText() As Document
			Dim doc As New Document()

			'Remove the previous changes of the document
			doc.RemoveAllChildren()

			'Set the document author
			doc.BuiltInDocumentProperties.Author = "Test Author"

			'Create paragraph without run
			Dim builder As New DocumentBuilder(doc)
			builder.Writeln()

			Return doc
		End Function

		''' <summary>
		''' Create new document with text
		''' </summary>
		Friend Shared Function CreateDocumentFillWithDummyText() As Document
			Dim doc As New Document()

			'Remove the previous changes of the document
			doc.RemoveAllChildren()

			'Set the document author
			doc.BuiltInDocumentProperties.Author = "Test Author"

			Dim builder As New DocumentBuilder(doc)

			builder.Write("Page ")
			builder.InsertField("PAGE", "")
			builder.Write(" of ")
			builder.InsertField("NUMPAGES", "")

			'Insert new table with two rows and two cells
			InsertTable(builder)

			builder.Writeln("Hello World!")

			' Continued on page 2 of the document content
			builder.InsertBreak(BreakType.PageBreak)

			'Insert TOC entries
			InsertToc(builder)

			Return doc
		End Function

		''' <summary>
		''' Create new document with textbox shape and some query
		''' </summary>
		Friend Shared Function CreateTemplateDocumentWithDrawObjects(ByVal templateText As String, ByVal shapeType As ShapeType) As Document
			Dim doc As New Document()

			' Create textbox shape.
			Dim shape As New Shape(doc, shapeType)
			shape.Width = 431.5
			shape.Height = 346.35

			Dim paragraph As New Paragraph(doc)
			paragraph.AppendChild(New Run(doc, templateText))

			' Insert paragraph into the textbox.
			shape.AppendChild(paragraph)

			' Insert textbox into the document.
			doc.FirstSection.Body.FirstParagraph.AppendChild(shape)

			Return doc
		End Function

		''' <summary>
		''' Insert new table in the document
		''' </summary>
		Private Shared Sub InsertTable(ByVal builder As DocumentBuilder)
			'Start creating a new table
			Dim table As Table = builder.StartTable()

			'Insert Row 1 Cell 1
			builder.InsertCell()
			builder.Write("Date")

			'Set width to fit the table contents
			table.AutoFit(AutoFitBehavior.AutoFitToContents)

			'Insert Row 1 Cell 2
			builder.InsertCell()
			builder.Write(" ")

			builder.EndRow()

			'Insert Row 2 Cell 1
			builder.InsertCell()
			builder.Write("Author")

			'Insert Row 2 Cell 2
			builder.InsertCell()
			builder.Write(" ")

			builder.EndRow()

			builder.EndTable()
		End Sub

		''' <summary>
		''' Insert TOC entries in the document
		''' </summary>
		Private Shared Sub InsertToc(ByVal builder As DocumentBuilder)
			' Creating TOC entries
			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1

			builder.Writeln("Heading 1")

			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2

			builder.Writeln("Heading 1.1")

			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4

			builder.Writeln("Heading 1.1.1.1")

			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5

			builder.Writeln("Heading 1.1.1.1.1")

			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading9

			builder.Writeln("Heading 1.1.1.1.1.1.1.1.1")
		End Sub

		''' <summary>
		''' Insert run into the current document
		''' </summary>
		''' <param name="doc">
		''' Current document
		''' </param>
		''' <param name="text">
		''' Custom text
		''' </param>
		''' <param name="paraIndex">
		''' Paragraph index
		''' </param>
		Friend Shared Function InsertNewRun(ByVal doc As Document, ByVal text As String, ByVal paraIndex As Integer) As Run
			Dim para As Paragraph = GetParagraph(doc, paraIndex)

			Dim run As New Run(doc) With {.Text = text}

			para.AppendChild(run)

			Return run
		End Function

		''' <summary>
		''' Insert text into the current document
		''' </summary>
		''' <param name="builder">
		''' Current document builder
		''' </param>
		''' <param name="textStrings">
		''' Custom text
		''' </param>
		Friend Shared Sub InsertBuilderText(ByVal builder As DocumentBuilder, ByVal textStrings() As String)
			For Each textString As String In textStrings
				builder.Writeln(textString)
			Next textString
		End Sub

		''' <summary>
		''' Get paragraph text of the current document
		''' </summary>
		''' <param name="doc">
		''' Current document
		''' </param>
		''' <param name="paraIndex">
		''' Paragraph number from collection
		''' </param>
		Friend Shared Function GetParagraphText(ByVal doc As Document, ByVal paraIndex As Integer) As String
			Return doc.FirstSection.Body.Paragraphs(paraIndex).GetText()
		End Function

		''' <summary>
		''' Get paragraph of the current document
		''' </summary>
		''' <param name="doc">
		''' Current document
		''' </param>
		''' <param name="paraIndex">
		''' Paragraph number from collection
		''' </param>
		Friend Shared Function GetParagraph(ByVal doc As Document, ByVal paraIndex As Integer) As Paragraph
			Return doc.FirstSection.Body.Paragraphs(paraIndex)
		End Function
	End Class
End Namespace
