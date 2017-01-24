' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExBorder
		Inherits ApiExampleBase
		<Test> _
		Public Sub FontBorder()
			'ExStart
			'ExFor:Border
			'ExFor:Border.Color
			'ExFor:Border.LineWidth
			'ExFor:Border.LineStyle
			'ExFor:Font.Border
			'ExFor:LineStyle
			'ExFor:Font
			'ExFor:DocumentBuilder.Font
			'ExFor:DocumentBuilder.Write
			'ExSummary:Inserts a string surrounded by a border into a document.
			Dim builder As New DocumentBuilder()

			builder.Font.Border.Color = Color.Green
			builder.Font.Border.LineWidth = 2.5
			builder.Font.Border.LineStyle = LineStyle.DashDotStroker

			builder.Write("run of text in a green border")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ParagraphTopBorder()
			'ExStart
			'ExFor:BorderCollection
			'ExFor:Border
			'ExFor:BorderType
			'ExFor:DocumentBuilder.ParagraphFormat
			'ExFor:DocumentBuilder.Writeln(String)
			'ExSummary:Inserts a paragraph with a top border.
			Dim builder As New DocumentBuilder()

			Dim topBorder As Border = builder.ParagraphFormat.Borders(BorderType.Top)
			topBorder.Color = Color.Red
			topBorder.LineStyle = LineStyle.DashSmallGap
			topBorder.LineWidth = 4

			builder.Writeln("Hello World!")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ClearFormattingEx()
			'ExStart
			'ExFor:Border.ClearFormatting
			'ExSummary:Shows how to remove borders from a paragraph one by one.
			Dim doc As New Document(MyDir & "Document.Borders.doc")
			Dim builder As New DocumentBuilder(doc)
			Dim borders As BorderCollection = builder.ParagraphFormat.Borders

			For Each border As Border In borders
				border.ClearFormatting()
			Next border

			builder.CurrentParagraph.Runs(0).Text = "Paragraph with no border"
			doc.Save(MyDir & "\Artifacts\Document.NoBorder.doc")
			'ExEnd
		End Sub
	End Class
End Namespace