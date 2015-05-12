'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.IO
Imports System.Reflection

Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Fields

Namespace AddWatermarkExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim doc As New Document(dataDir & "TestFile.doc")
			InsertWatermarkText(doc, "CONFIDENTIAL")
			doc.Save(dataDir & "TestFile Out.doc")
		End Sub

		''' <summary>
		''' Inserts a watermark into a document.
		''' </summary>
		''' <param name="doc">The input document.</param>
		''' <param name="watermarkText">Text of the watermark.</param>
		Private Shared Sub InsertWatermarkText(ByVal doc As Document, ByVal watermarkText As String)
			' Create a watermark shape. This will be a WordArt shape. 
			' You are free to try other shape types as watermarks.
			Dim watermark As New Shape(doc, ShapeType.TextPlainText)

			' Set up the text of the watermark.
			watermark.TextPath.Text = watermarkText
			watermark.TextPath.FontFamily = "Arial"
			watermark.Width = 500
			watermark.Height = 100
			' Text will be directed from the bottom-left to the top-right corner.
			watermark.Rotation = -40
			' Remove the following two lines if you need a solid black text.
			watermark.Fill.Color = Color.Gray ' Try LightGray to get more Word-style watermark
			watermark.StrokeColor = Color.Gray ' Try LightGray to get more Word-style watermark

			' Place the watermark in the page center.
			watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page
			watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page
			watermark.WrapType = WrapType.None
			watermark.VerticalAlignment = VerticalAlignment.Center
			watermark.HorizontalAlignment = HorizontalAlignment.Center

			' Create a new paragraph and append the watermark to this paragraph.
			Dim watermarkPara As New Paragraph(doc)
			watermarkPara.AppendChild(watermark)

			' Insert the watermark into all headers of each document section.
			For Each sect As Section In doc.Sections
				' There could be up to three different headers in each section, since we want
				' the watermark to appear on all pages, insert into all headers.
				InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary)
				InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst)
				InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven)
			Next sect
		End Sub

		Private Shared Sub InsertWatermarkIntoHeader(ByVal watermarkPara As Paragraph, ByVal sect As Section, ByVal headerType As HeaderFooterType)
			Dim header As HeaderFooter = sect.HeadersFooters(headerType)

			If header Is Nothing Then
				' There is no header of the specified type in the current section, create it.
				header = New HeaderFooter(sect.Document, headerType)
				sect.HeadersFooters.Add(header)
			End If

			' Insert a clone of the watermark into the header.
			header.AppendChild(watermarkPara.Clone(True))
		End Sub
	End Class
End Namespace
'ExEnd