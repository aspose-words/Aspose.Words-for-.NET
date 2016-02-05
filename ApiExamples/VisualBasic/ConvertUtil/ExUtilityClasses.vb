' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework


Namespace ApiExamples.ConvertUtil
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
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As Aspose.Words.PageSetup = builder.PageSetup
			pageSetup.TopMargin = Aspose.Words.ConvertUtil.InchToPoint(1.0)
			pageSetup.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(1.0)
			pageSetup.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(1.5)
			pageSetup.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(1.5)
			pageSetup.HeaderDistance = Aspose.Words.ConvertUtil.InchToPoint(0.2)
			pageSetup.FooterDistance = Aspose.Words.ConvertUtil.InchToPoint(0.2)
			'ExEnd
		End Sub

		<Test> _
		Public Sub MillimeterToPointEx()
			'ExStart
			'ExFor:ConvertUtil.MillimeterToPoint
			'ExSummary:Shows how to specify page properties in millimeters.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As Aspose.Words.PageSetup = builder.PageSetup
			pageSetup.TopMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(25.0)
			pageSetup.BottomMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(25.0)
			pageSetup.LeftMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(37.5)
			pageSetup.RightMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(37.5)
			pageSetup.HeaderDistance = Aspose.Words.ConvertUtil.MillimeterToPoint(5.0)
			pageSetup.FooterDistance = Aspose.Words.ConvertUtil.MillimeterToPoint(5.0)

			builder.Writeln("Hello world.")
			builder.Document.Save(MyDir & "PageSetup.PageMargins Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub PointToInchEx()
			'ExStart
			'ExFor:ConvertUtil.PointToInch
			'ExSummary:Shows how to convert points to inches.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As Aspose.Words.PageSetup = builder.PageSetup
			pageSetup.TopMargin = Aspose.Words.ConvertUtil.InchToPoint(2.0)

			Console.WriteLine("The size of my top margin is {0} points, or {1} inches.", pageSetup.TopMargin, Aspose.Words.ConvertUtil.PointToInch(pageSetup.TopMargin))
			'ExEnd
		End Sub

		<Test> _
		Public Sub PixelToPointEx()
			'ExStart
			'ExFor:ConvertUtil.PixelToPoint(double)
			'ExFor:ConvertUtil.PixelToPoint(double, double)
			'ExSummary:Shows how to specify page properties in pixels with default and custom resolution.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetupNoDpi As Aspose.Words.PageSetup = builder.PageSetup
			pageSetupNoDpi.TopMargin = Aspose.Words.ConvertUtil.PixelToPoint(100.0)
			pageSetupNoDpi.BottomMargin = Aspose.Words.ConvertUtil.PixelToPoint(100.0)
			pageSetupNoDpi.LeftMargin = Aspose.Words.ConvertUtil.PixelToPoint(150.0)
			pageSetupNoDpi.RightMargin = Aspose.Words.ConvertUtil.PixelToPoint(150.0)
			pageSetupNoDpi.HeaderDistance = Aspose.Words.ConvertUtil.PixelToPoint(20.0)
			pageSetupNoDpi.FooterDistance = Aspose.Words.ConvertUtil.PixelToPoint(20.0)

			builder.Writeln("Hello world.")
			builder.Document.Save(MyDir & "PageSetup.PageMargins.DefaultResolution Out.doc")

			Dim myDpi As Double = 150.0

			Dim pageSetupWithDpi As Aspose.Words.PageSetup = builder.PageSetup
			pageSetupWithDpi.TopMargin = Aspose.Words.ConvertUtil.PixelToPoint(100.0, myDpi)
			pageSetupWithDpi.BottomMargin = Aspose.Words.ConvertUtil.PixelToPoint(100.0, myDpi)
			pageSetupWithDpi.LeftMargin = Aspose.Words.ConvertUtil.PixelToPoint(150.0, myDpi)
			pageSetupWithDpi.RightMargin = Aspose.Words.ConvertUtil.PixelToPoint(150.0, myDpi)
			pageSetupWithDpi.HeaderDistance = Aspose.Words.ConvertUtil.PixelToPoint(20.0, myDpi)
			pageSetupWithDpi.FooterDistance = Aspose.Words.ConvertUtil.PixelToPoint(20.0, myDpi)

			builder.Document.Save(MyDir & "PageSetup.PageMargins.CustomResolution Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub PointToPixelEx()
			'ExStart
			'ExFor:ConvertUtil.PointToPixel(double)
			'ExFor:ConvertUtil.PointToPixel(double, double)
			'ExSummary:Shows how to use convert points to pixels with default and custom resolution.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As Aspose.Words.PageSetup = builder.PageSetup
			pageSetup.TopMargin = Aspose.Words.ConvertUtil.PixelToPoint(2.0)

			Dim myDpi As Double = 192.0

			Console.WriteLine("The size of my top margin is {0} points, or {1} pixels with default resolution.", pageSetup.TopMargin, Aspose.Words.ConvertUtil.PointToPixel(pageSetup.TopMargin))

			Console.WriteLine("The size of my top margin is {0} points, or {1} pixels with custom resolution.", pageSetup.TopMargin, Aspose.Words.ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi))
			'ExEnd
		End Sub

		<Test> _
		Public Sub PixelToNewDpiEx()
			'ExStart
			'ExFor:ConvertUtil.PixelToNewDpi
			'ExSummary:Shows how to check how an amount of pixels changes when the dpi is changed.
			Dim doc As New Aspose.Words.Document()
			Dim builder As New DocumentBuilder(doc)

			Dim pageSetup As Aspose.Words.PageSetup = builder.PageSetup
			pageSetup.TopMargin = 72
			Dim oldDpi As Double = 92.0
			Dim newDpi As Double = 192.0

			Console.WriteLine("{0} pixels at {1} dpi becomes {2} pixels at {3} dpi.", pageSetup.TopMargin, oldDpi, Aspose.Words.ConvertUtil.PixelToNewDpi(pageSetup.TopMargin, oldDpi, newDpi), newDpi)
			'ExEnd
		End Sub
	End Class
End Namespace
