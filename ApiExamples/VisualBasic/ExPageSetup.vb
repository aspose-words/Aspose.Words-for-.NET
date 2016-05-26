' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.Drawing.Printing

Imports Aspose.Words

Imports NUnit.Framework

Imports PaperSize = Aspose.Words.PaperSize

Namespace ApiExamples
	<TestFixture> _
	Public Class ExPageSetup
		Inherits ApiExampleBase
		<Test> _
		Public Sub ClearFormatting()
			'ExStart
			'ExFor:DocumentBuilder.PageSetup
			'ExFor:DocumentBuilder.InsertBreak
			'ExFor:DocumentBuilder.Document
			'ExFor:PageSetup
			'ExFor:PageSetup.Orientation
			'ExFor:PageSetup.VerticalAlignment
			'ExFor:PageSetup.ClearFormatting
			'ExFor:Orientation
			'ExFor:PageVerticalAlignment
			'ExFor:BreakType
			'ExSummary:Shows how to insert sections using DocumentBuilder, specify page setup for a section and reset page setup to defaults.
			Dim builder As New DocumentBuilder()

			' Modify the first section in the document.
			builder.PageSetup.Orientation = Orientation.Landscape
			builder.PageSetup.VerticalAlignment = PageVerticalAlignment.Center
			builder.Writeln("Section 1, landscape oriented and text vertically centered.")

			' Start a new section and reset its formatting to defaults.
			builder.InsertBreak(BreakType.SectionBreakNewPage)
			builder.PageSetup.ClearFormatting()
			builder.Writeln("Section 2, back to default Letter paper size, portrait orientation and top alignment.")

			builder.Document.Save(MyDir & "\Artifacts\PageSetup.ClearFormatting.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub DifferentHeaders()
			'ExStart
			'ExFor:PageSetup.DifferentFirstPageHeaderFooter
			'ExFor:PageSetup.OddAndEvenPagesHeaderFooter
			'ExSummary:Creates headers and footers different for first, even and odd pages using DocumentBuilder.
			Dim builder As New DocumentBuilder()

			Dim ps As PageSetup = builder.PageSetup
			ps.DifferentFirstPageHeaderFooter = True
			ps.OddAndEvenPagesHeaderFooter = True

			builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst)
			builder.Writeln("First page header.")

			builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven)
			builder.Writeln("Even pages header.")

			builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary)
			builder.Writeln("Odd pages header.")

			' Move back to the main story of the first section.
			builder.MoveToSection(0)
			builder.Writeln("Text page 1.")
			builder.InsertBreak(BreakType.PageBreak)
			builder.Writeln("Text page 2.")
			builder.InsertBreak(BreakType.PageBreak)
			builder.Writeln("Text page 3.")

			builder.Document.Save(MyDir & "\Artifacts\PageSetup.DifferentHeaders.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub SectionStart()
			'ExStart
			'ExFor:SectionStart
			'ExFor:PageSetup.SectionStart
			'ExFor:Document.Sections
			'ExSummary:Specifies how the section starts, from a new page, on the same page or other.
			Dim doc As New Document()
			doc.Sections(0).PageSetup.SectionStart = Aspose.Words.SectionStart.Continuous
			'ExEnd
		End Sub

		<Test> _
		Public Sub DefaultPaperTray()
			'ExStart
			'ExFor:PageSetup.FirstPageTray
			'ExFor:PageSetup.OtherPagesTray
			'ExSummary:Changes all sections in a document to use the default paper tray of the selected printer.
			Dim doc As New Document()

			' Find the printer that will be used for printing this document. In this case it is the default printer.
			' You can define a specific printer using PrinterName.
			Dim settings As New PrinterSettings()

			' The paper tray value stored in documents is completely printer specific. This means 
			' The code below resets all page tray values to use the current printers default tray.
			' You can enumerate PrinterSettings.PaperSources to find the other valid paper tray values of the selected printer.
			For Each section As Section In doc.Sections
				section.PageSetup.FirstPageTray = settings.DefaultPageSettings.PaperSource.RawKind
				section.PageSetup.OtherPagesTray = settings.DefaultPageSettings.PaperSource.RawKind
			Next section
			'ExEnd
		End Sub

		<Test, Explicit> _
		Public Sub PaperTrayForDifferentPaperType()
			'ExStart
			'ExFor:PageSetup.FirstPageTray
			'ExFor:PageSetup.OtherPagesTray
			'ExSummary:Shows how to set up printing using different printer trays for different paper sizes.
			Dim doc As New Document()

			' Choose the default printer to be used for printing this document.
			Dim settings As New PrinterSettings()

			' This is the tray we will use for A4 paper size. This is the first tray in the paper sources collection.
			Dim printerTrayForA4 As Integer = settings.PaperSources(0).RawKind
			' The is the tray we will use for Letter paper size. This is the second tray in the paper sources collection.
			Dim printerTrayForLetter As Integer = settings.PaperSources(1).RawKind

			' Set the page tray used for each section based off the paper size used in the section.
			For Each section As Section In doc.Sections
                If section.PageSetup.PaperSize = PaperSize.Letter Then
                    section.PageSetup.FirstPageTray = printerTrayForLetter
                    section.PageSetup.OtherPagesTray = printerTrayForLetter
                ElseIf section.PageSetup.PaperSize = PaperSize.A4 Then
                    section.PageSetup.FirstPageTray = printerTrayForA4
                    section.PageSetup.OtherPagesTray = printerTrayForA4
                End If
			Next section
			'ExEnd
		End Sub

		<Test> _
		Public Sub PageMargins()
			'ExStart
			'ExFor:ConvertUtil
			'ExFor:ConvertUtil.InchToPoint
			'ExFor:PaperSize
			'ExFor:PageSetup.PaperSize
			'ExFor:PageSetup.Orientation
			'ExFor:PageSetup.TopMargin
			'ExFor:PageSetup.BottomMargin
			'ExFor:PageSetup.LeftMargin
			'ExFor:PageSetup.RightMargin
			'ExFor:PageSetup.HeaderDistance
			'ExFor:PageSetup.FooterDistance
			'ExSummary:Specifies paper size, orientation, margins and other settings for a section.
			Dim builder As New DocumentBuilder()

			Dim ps As PageSetup = builder.PageSetup
			ps.PaperSize = PaperSize.Legal
			ps.Orientation = Orientation.Landscape
			ps.TopMargin = ConvertUtil.InchToPoint(1.0)
			ps.BottomMargin = ConvertUtil.InchToPoint(1.0)
			ps.LeftMargin = ConvertUtil.InchToPoint(1.5)
			ps.RightMargin = ConvertUtil.InchToPoint(1.5)
			ps.HeaderDistance = ConvertUtil.InchToPoint(0.2)
			ps.FooterDistance = ConvertUtil.InchToPoint(0.2)

			builder.Writeln("Hello world.")

			builder.Document.Save(MyDir & "\Artifacts\PageSetup.PageMargins.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ColumnsSameWidth()
			'ExStart
			'ExFor:PageSetup.TextColumns
			'ExFor:TextColumnCollection
			'ExFor:TextColumnCollection.Spacing
			'ExFor:TextColumnCollection.SetCount
			'ExSummary:Creates multiple evenly spaced columns in a section using DocumentBuilder.
			Dim builder As New DocumentBuilder()

			Dim columns As TextColumnCollection = builder.PageSetup.TextColumns
			' Make spacing between columns wider.
			columns.Spacing = 100
			' This creates two columns of equal width.
			columns.SetCount(2)

			builder.Writeln("Text in column 1.")
			builder.InsertBreak(BreakType.ColumnBreak)
			builder.Writeln("Text in column 2.")

			builder.Document.Save(MyDir & "\Artifacts\PageSetup.ColumnsSameWidth.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ColumnsCustomWidth()
			'ExStart
			'ExFor:TextColumnCollection.LineBetween
			'ExFor:TextColumnCollection.EvenlySpaced
			'ExFor:TextColumnCollection.Item
			'ExFor:TextColumn
			'ExFor:TextColumn.Width
			'ExFor:TextColumn.SpaceAfter
			'ExSummary:Creates multiple columns of different widths in a section using DocumentBuilder.
			Dim builder As New DocumentBuilder()

			Dim columns As TextColumnCollection = builder.PageSetup.TextColumns
			' Show vertical line between columns.
			columns.LineBetween = True
			' Indicate we want to create column with different widths.
			columns.EvenlySpaced = False
			' Create two columns, note they will be created with zero widths, need to set them.
			columns.SetCount(2)

			' Set the first column to be narrow.
			Dim c1 As TextColumn = columns(0)
			c1.Width = 100
			c1.SpaceAfter = 20

			' Set the second column to take the rest of the space available on the page.
			Dim c2 As TextColumn = columns(1)
			Dim ps As PageSetup = builder.PageSetup
			Dim contentWidth As Double = ps.PageWidth - ps.LeftMargin - ps.RightMargin
			c2.Width = contentWidth - c1.Width - c1.SpaceAfter

			builder.Writeln("Narrow column 1.")
			builder.InsertBreak(BreakType.ColumnBreak)
			builder.Writeln("Wide column 2.")

			builder.Document.Save(MyDir & "\Artifacts\PageSetup.ColumnsCustomWidth.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub LineNumbers()
			'ExStart
			'ExFor:PageSetup.LineStartingNumber
			'ExFor:PageSetup.LineNumberCountBy
			'ExFor:PageSetup.LineNumberRestartMode
			'ExFor:LineNumberRestartMode
			'ExSummary:Turns on Microsoft Word line numbering for a section.
			Dim builder As New DocumentBuilder()

			Dim ps As PageSetup = builder.PageSetup
			ps.LineStartingNumber = 1
			ps.LineNumberCountBy = 5
			ps.LineNumberRestartMode = LineNumberRestartMode.RestartPage

			For i As Integer = 1 To 20
				builder.Writeln(String.Format("Line {0}.", i))
			Next i

			builder.Document.Save(MyDir & "\Artifacts\PageSetup.LineNumbers.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub PageBorderProperties()
			'ExStart
			'ExFor:Section.PageSetup
			'ExFor:PageSetup.BorderAlwaysInFront
			'ExFor:PageSetup.BorderDistanceFrom
			'ExFor:PageSetup.BorderAppliesTo
			'ExFor:PageBorderDistanceFrom
			'ExFor:PageBorderAppliesTo
			'ExFor:Border.DistanceFromText
			'ExSummary:Creates a page border that looks like a wide blue band at the top of the first page only.
			Dim doc As New Document()

			Dim ps As PageSetup = doc.Sections(0).PageSetup
			ps.BorderAlwaysInFront = False
			ps.BorderDistanceFrom = PageBorderDistanceFrom.PageEdge
			ps.BorderAppliesTo = PageBorderAppliesTo.FirstPage

			Dim border As Border = ps.Borders(BorderType.Top)
			border.LineStyle = LineStyle.Single
			border.LineWidth = 30
			border.Color = Color.Blue
			border.DistanceFromText = 0

			doc.Save(MyDir & "\Artifacts\PageSetup.PageBorderTop.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub PageBorders()
			'ExStart
			'ExFor:PageSetup.Borders
			'ExFor:Border.Shadow
			'ExFor:BorderCollection.LineStyle
			'ExFor:BorderCollection.LineWidth
			'ExFor:BorderCollection.Color
			'ExFor:BorderCollection.DistanceFromText
			'ExFor:BorderCollection.Shadow
			'ExSummary:Creates a fancy looking green wavy page border with a shadow.
			Dim doc As New Document()
			Dim ps As PageSetup = doc.Sections(0).PageSetup

			ps.Borders.LineStyle = LineStyle.DoubleWave
			ps.Borders.LineWidth = 2
			ps.Borders.Color = Color.Green
			ps.Borders.DistanceFromText = 24
			ps.Borders.Shadow = True

			doc.Save(MyDir & "\Artifacts\PageSetup.PageBorders.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub PageNumbering()
			'ExStart
			'ExFor:PageSetup.RestartPageNumbering
			'ExFor:PageSetup.PageStartingNumber
			'ExFor:PageSetup.PageNumberStyle
			'ExFor:DocumentBuilder.InsertField(string, string)
			'ExSummary:Shows how to control page numbering per section.
			' This document has two sections, but no page numbers yet.
			Dim doc As New Document(MyDir & "PageSetup.PageNumbering.doc")

			' Use document builder to create a header with a page number field for the first section.
			' The page number will look like "Page V".
			Dim builder As New DocumentBuilder(doc)
			builder.MoveToSection(0)
			builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary)
			builder.Write("Page ")
			builder.InsertField("PAGE", "")

			' Set first section page numbering.
			Dim section As Section = doc.Sections(0)
			section.PageSetup.RestartPageNumbering = True
			section.PageSetup.PageStartingNumber = 5
			section.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman


			' Create a header for the section section. 
			' The page number will look like " - 10 - ".
			builder.MoveToSection(1)
			builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary)
			builder.ParagraphFormat.Alignment = ParagraphAlignment.Center
			builder.Write(" - ")
			builder.InsertField("PAGE", "")
			builder.Write(" - ")

			' Set second section page numbering.
			section = doc.Sections(1)
			section.PageSetup.PageStartingNumber = 10
			section.PageSetup.RestartPageNumbering = True
			section.PageSetup.PageNumberStyle = NumberStyle.Arabic

			doc.Save(MyDir & "\Artifacts\PageSetup.PageNumbering.doc")
			'ExEnd
		End Sub
	End Class
End Namespace
