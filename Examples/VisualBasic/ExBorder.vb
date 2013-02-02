'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework

Namespace Examples
	<TestFixture> _
	Public Class ExBorder
		Inherits ExBase
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

			builder.Font.Border.Color = System.Drawing.Color.Green
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
			topBorder.Color = System.Drawing.Color.Red
			topBorder.LineStyle = LineStyle.DashSmallGap
			topBorder.LineWidth = 4

			builder.Writeln("Hello World!")
			'ExEnd
		End Sub
	End Class
End Namespace
