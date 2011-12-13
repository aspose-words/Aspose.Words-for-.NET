'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
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
	Public Class ExUtilityClasses
		Inherits ExBase
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
	End Class
End Namespace
