' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExUtilityClasses
		Inherits ApiExampleBase
		<Test> _
		Public Sub UtilityClassesUseControlCharacters()
			Dim text As String = "test" & Constants.vbCr
			'ExStart
			'ExFor:ControlChar
			'ExFor:ControlChar.Cr
			'ExFor:ControlChar.CrLf
			'ExId:UtilityClassesUseControlCharacters
			'ExSummary:Shows how to use control characters.
			' Replace "\r" control character with "\r\n"
			text = text.Replace(ControlChar.Cr, ControlChar.CrLf)
			'ExEnd
		End Sub

		<Test> _
		Public Sub UtilityClassesConvertBetweenMeasurementUnits()
			'ExStart
			'ExFor:ConvertUtil
			'ExId:UtilityClassesConvertBetweenMeasurementUnits
			'ExSummary:Shows how to specify page properties in inches.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As PageSetup = builder.PageSetup
			pageSetup.TopMargin = ConvertUtil.InchToPoint(1.0)
			pageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0)
			pageSetup.LeftMargin = ConvertUtil.InchToPoint(1.5)
			pageSetup.RightMargin = ConvertUtil.InchToPoint(1.5)
			pageSetup.HeaderDistance = ConvertUtil.InchToPoint(0.2)
			pageSetup.FooterDistance = ConvertUtil.InchToPoint(0.2)
			'ExEnd
		End Sub

		<Test> _
		Public Sub MillimeterToPointEx()
			'ExStart
			'ExFor:ConvertUtil.MillimeterToPoint
			'ExSummary:Shows how to specify page properties in millimeters.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As PageSetup = builder.PageSetup
			pageSetup.TopMargin = ConvertUtil.MillimeterToPoint(25.0)
			pageSetup.BottomMargin = ConvertUtil.MillimeterToPoint(25.0)
			pageSetup.LeftMargin = ConvertUtil.MillimeterToPoint(37.5)
			pageSetup.RightMargin = ConvertUtil.MillimeterToPoint(37.5)
			pageSetup.HeaderDistance = ConvertUtil.MillimeterToPoint(5.0)
			pageSetup.FooterDistance = ConvertUtil.MillimeterToPoint(5.0)

			builder.Writeln("Hello world.")
			builder.Document.Save(MyDir & "\Artifacts\PageSetup.PageMargins.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub PointToInchEx()
			'ExStart
			'ExFor:ConvertUtil.PointToInch
			'ExSummary:Shows how to convert points to inches.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As PageSetup = builder.PageSetup
			pageSetup.TopMargin = ConvertUtil.InchToPoint(2.0)

			Console.WriteLine("The size of my top margin is {0} points, or {1} inches.", pageSetup.TopMargin, ConvertUtil.PointToInch(pageSetup.TopMargin))
			'ExEnd
		End Sub

		<Test> _
		Public Sub PixelToPointEx()
			'ExStart
			'ExFor:ConvertUtil.PixelToPoint(double)
			'ExFor:ConvertUtil.PixelToPoint(double, double)
			'ExSummary:Shows how to specify page properties in pixels with default and custom resolution.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetupNoDpi As PageSetup = builder.PageSetup
			pageSetupNoDpi.TopMargin = ConvertUtil.PixelToPoint(100.0)
			pageSetupNoDpi.BottomMargin = ConvertUtil.PixelToPoint(100.0)
			pageSetupNoDpi.LeftMargin = ConvertUtil.PixelToPoint(150.0)
			pageSetupNoDpi.RightMargin = ConvertUtil.PixelToPoint(150.0)
			pageSetupNoDpi.HeaderDistance = ConvertUtil.PixelToPoint(20.0)
			pageSetupNoDpi.FooterDistance = ConvertUtil.PixelToPoint(20.0)

			builder.Writeln("Hello world.")
			builder.Document.Save(MyDir & "\Artifacts\PageSetup.PageMargins.DefaultResolution.doc")

			Dim myDpi As Double = 150.0

			Dim pageSetupWithDpi As PageSetup = builder.PageSetup
			pageSetupWithDpi.TopMargin = ConvertUtil.PixelToPoint(100.0, myDpi)
			pageSetupWithDpi.BottomMargin = ConvertUtil.PixelToPoint(100.0, myDpi)
			pageSetupWithDpi.LeftMargin = ConvertUtil.PixelToPoint(150.0, myDpi)
			pageSetupWithDpi.RightMargin = ConvertUtil.PixelToPoint(150.0, myDpi)
			pageSetupWithDpi.HeaderDistance = ConvertUtil.PixelToPoint(20.0, myDpi)
			pageSetupWithDpi.FooterDistance = ConvertUtil.PixelToPoint(20.0, myDpi)

			builder.Document.Save(MyDir & "\Artifacts\PageSetup.PageMargins.CustomResolution.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub PointToPixelEx()
			'ExStart
			'ExFor:ConvertUtil.PointToPixel(double)
			'ExFor:ConvertUtil.PointToPixel(double, double)
			'ExSummary:Shows how to use convert points to pixels with default and custom resolution.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As PageSetup = builder.PageSetup
			pageSetup.TopMargin = ConvertUtil.PixelToPoint(2.0)

			Dim myDpi As Double = 192.0

			Console.WriteLine("The size of my top margin is {0} points, or {1} pixels with default resolution.", pageSetup.TopMargin, ConvertUtil.PointToPixel(pageSetup.TopMargin))

			Console.WriteLine("The size of my top margin is {0} points, or {1} pixels with custom resolution.", pageSetup.TopMargin, ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi))
			'ExEnd
		End Sub

		<Test> _
		Public Sub PixelToNewDpiEx()
			'ExStart
			'ExFor:ConvertUtil.PixelToNewDpi
			'ExSummary:Shows how to check how an amount of pixels changes when the dpi is changed.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As PageSetup = builder.PageSetup
			pageSetup.TopMargin = 72
			Dim oldDpi As Double = 92.0
			Dim newDpi As Double = 192.0

			Console.WriteLine("{0} pixels at {1} dpi becomes {2} pixels at {3} dpi.", pageSetup.TopMargin, oldDpi, ConvertUtil.PixelToNewDpi(pageSetup.TopMargin, oldDpi, newDpi), newDpi)
			'ExEnd
		End Sub
	End Class
End Namespace
