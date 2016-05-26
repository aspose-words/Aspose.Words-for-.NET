' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports Aspose.Words.Saving

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Friend Class ExHtmlSaveOptions
		Inherits ApiExampleBase
		'For assert this test you need to open html docs and they shouldn't have negative left margins
		<Test, TestCase(SaveFormat.Html), TestCase(SaveFormat.Mhtml), TestCase(SaveFormat.Epub)> _
		Public Sub ExportPageMargins(ByVal saveFormat As SaveFormat)
			Dim doc As New Document(MyDir & "HtmlSaveOptions.ExportPageMargins.docx")

			Dim htmlSaveOptions As HtmlSaveOptions = New HtmlSaveOptions With {.SaveFormat = saveFormat, .ExportPageMargins = True}

			Select Case saveFormat
				Case SaveFormat.Html
					doc.Save(MyDir & "ExportPageMargins.html", htmlSaveOptions)
				Case SaveFormat.Mhtml
					doc.Save(MyDir & "ExportPageMargins.Mhtml", htmlSaveOptions)
				Case SaveFormat.Epub
					doc.Save(MyDir & "ExportPageMargins.Epub", htmlSaveOptions) 'There is draw images bug with epub. Need write to NSezganov
			End Select
		End Sub
	End Class
End Namespace
