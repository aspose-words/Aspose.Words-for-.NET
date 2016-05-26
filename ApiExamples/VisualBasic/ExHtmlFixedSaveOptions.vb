' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System.Text

Imports Aspose.Words
Imports Aspose.Words.Saving

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Friend Class ExHtmlFixedSaveOptions
		Inherits ApiExampleBase
		<Test> _
		Public Sub UseEncoding()
			'ExStart
			'ExFor:Saving.HtmlFixedSaveOptions.Encoding
			'ExSummary:Shows how to use "Encoding" parameter with "HtmlFixedSaveOptions"
			Dim doc As New Document()

			Dim builder As New DocumentBuilder(doc)
			builder.Writeln("Hello World!")

			'Create "HtmlFixedSaveOptions" with "Encoding" parameter
			'You can also set "Encoding" using System.Text.Encoding, like "Encoding.ASCII", or "Encoding.GetEncoding()"
			Dim htmlFixedSaveOptions As HtmlFixedSaveOptions = New HtmlFixedSaveOptions With {.Encoding = New ASCIIEncoding(), .SaveFormat = SaveFormat.HtmlFixed}

			'Uses "HtmlFixedSaveOptions"
			doc.Save(MyDir & "\Artifacts\UseEncoding.html", htmlFixedSaveOptions)
			'ExEnd
		End Sub

		'Note: Tests doesn't containt validation result, because it's may take a lot of time for assert result
		'For validation result, you can save the document to html file and check out with notepad++, that file encoding will be correctly displayed (Encoding tab in Notepad++)
		<Test> _
		Public Sub EncodingUsingSystemTextEncoding()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			Dim htmlFixedSaveOptions As HtmlFixedSaveOptions = New HtmlFixedSaveOptions With {.Encoding = Encoding.ASCII, .SaveFormat = SaveFormat.HtmlFixed, .ExportEmbeddedCss = True, .ExportEmbeddedFonts = True, .ExportEmbeddedImages = True, .ExportEmbeddedSvg = True}

			doc.Save(MyDir & "EncodingUsingSystemTextEncoding.html", htmlFixedSaveOptions)
		End Sub

		<Test> _
		Public Sub EncodingUsingNewEncoding()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			Dim htmlFixedSaveOptions As HtmlFixedSaveOptions = New HtmlFixedSaveOptions With {.Encoding = New UTF32Encoding(), .SaveFormat = SaveFormat.HtmlFixed, .ExportEmbeddedCss = True, .ExportEmbeddedFonts = True, .ExportEmbeddedImages = True, .ExportEmbeddedSvg = True}

			doc.Save(MyDir & "EncodingUsingNewEncoding.html", htmlFixedSaveOptions)
		End Sub

		<Test> _
		Public Sub EncodingUsingGetEncoding()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			Dim htmlFixedSaveOptions As HtmlFixedSaveOptions = New HtmlFixedSaveOptions With {.Encoding = Encoding.GetEncoding("utf-16"), .SaveFormat = SaveFormat.HtmlFixed, .ExportEmbeddedCss = True, .ExportEmbeddedFonts = True, .ExportEmbeddedImages = True, .ExportEmbeddedSvg = True}

			doc.Save(MyDir & "EncodingUsingGetEncoding.html", htmlFixedSaveOptions)
		End Sub

		<Test, TestCase(True), TestCase(False)> _
		Public Sub ExportFormFields(ByVal exportFormFields As Boolean)
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			builder.InsertCheckBox("CheckBox", False, 15)

			Dim htmlFixedSaveOptions As HtmlFixedSaveOptions = New HtmlFixedSaveOptions With {.SaveFormat = SaveFormat.HtmlFixed, .ExportEmbeddedCss = True, .ExportEmbeddedFonts = True, .ExportEmbeddedImages = True, .ExportEmbeddedSvg = True, .ExportFormFields = exportFormFields}

			'For assert test result you need to open documents and check that checkbox are clickable in "ExportFormFiels.html" file and are not clickable in "WithoutExportFormFiels.html" file
			If exportFormFields = True Then
				doc.Save(MyDir & "ExportFormFiels.html", htmlFixedSaveOptions)
			Else
				doc.Save(MyDir & "WithoutExportFormFiels.html", htmlFixedSaveOptions)
			End If
		End Sub
	End Class
End Namespace
