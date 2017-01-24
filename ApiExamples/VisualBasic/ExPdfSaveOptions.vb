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
Imports Aspose.Pdf.Facades
Imports Aspose.Pdf.Text

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Friend Class ExPdfSaveOptions
		Inherits ApiExampleBase
		<Test> _
		Public Sub CreateMissingOutlineLevels()
			'ExStart
			'ExFor:Saving.PdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels
			'ExSummary:Shows how to create missing outline levels saving the document in pdf
			Dim doc As New Document()

			Dim builder As New DocumentBuilder(doc)

			' Creating TOC entries
			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1

			builder.Writeln("Heading 1")

			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4

			builder.Writeln("Heading 1.1.1.1")
			builder.Writeln("Heading 1.1.1.2")

			builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading9

			builder.Writeln("Heading 1.1.1.1.1.1.1.1.1")
			builder.Writeln("Heading 1.1.1.1.1.1.1.1.2")

			'Create "PdfSaveOptions" with some mandatory parameters
			'"HeadingsOutlineLevels" specifies how many levels of headings to include in the document outline
			'"CreateMissingOutlineLevels" determining whether or not to create missing heading levels
			Dim pdfSaveOptions As New PdfSaveOptions()

			pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9
			pdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels = True
			pdfSaveOptions.SaveFormat = SaveFormat.Pdf

			doc.Save(MyDir & "\Artifacts\CreateMissingOutlineLevels.pdf", pdfSaveOptions)
			'ExEnd

			'Bind pdf with Aspose PDF
			Dim bookmarkEditor As New PdfBookmarkEditor()
			bookmarkEditor.BindPdf(MyDir & "\Artifacts\CreateMissingOutlineLevels.pdf")

			'Get all bookmarks from the document
			Dim bookmarks As Bookmarks = bookmarkEditor.ExtractBookmarks()

			Assert.AreEqual(11, bookmarks.Count)
		End Sub

		'Note: Test doesn't containt validation result, because it's difficult
		'For validation result, you can add some shapes to the document and assert, that the DML shapes are render correctly
		<Test> _
		Public Sub DrawingMl()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			Dim pdfSaveOptions As New PdfSaveOptions()
			pdfSaveOptions.DmlRenderingMode = DmlRenderingMode.DrawingML

			doc.Save(MyDir & "\Artifacts\DrawingMl.pdf", pdfSaveOptions)
		End Sub

		<Test> _
		Public Sub WithoutUpdateFields()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			Dim pdfSaveOptions As New PdfSaveOptions()
			pdfSaveOptions.UpdateFields = False

			doc.Save(MyDir & "\Artifacts\UpdateFields_False.pdf", pdfSaveOptions)

			Dim pdfDocument As New Aspose.Pdf.Document(MyDir & "\Artifacts\UpdateFields_False.pdf")

			'Get text fragment by search string
			Dim textFragmentAbsorber As New TextFragmentAbsorber("Page  of")
			pdfDocument.Pages.Accept(textFragmentAbsorber)

			'Assert that fields are not updated
			Assert.AreEqual("Page  of", textFragmentAbsorber.TextFragments(1).Text)
		End Sub

		<Test> _
		Public Sub WithUpdateFields()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			Dim pdfSaveOptions As New PdfSaveOptions()
			pdfSaveOptions.UpdateFields = True

			doc.Save(MyDir & "\Artifacts\UpdateFields_False.pdf", pdfSaveOptions)

			Dim pdfDocument As New Aspose.Pdf.Document(MyDir & "\Artifacts\UpdateFields_False.pdf")

			'Get text fragment by search string
			Dim textFragmentAbsorber As New TextFragmentAbsorber("Page 1 of 2")
			pdfDocument.Pages.Accept(textFragmentAbsorber)

			'Assert that fields are updated
			Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments(1).Text)
		End Sub

		'For assert this test you need to open "SaveOptions.PdfImageComppression PDF_A_1_B Out.pdf" and "SaveOptions.PdfImageComppression PDF_A_1_A Out.pdf" and check that header image in this documents are equal header image in the "SaveOptions.PdfImageComppression Out.pdf" 
		<Test> _
		Public Sub ImageCompression()
			'ExStart
			'ExFor:PdfSaveOptions.Compliance
			'ExFor:PdfSaveOptions.ImageCompression
			'ExFor:PdfSaveOptions.JpegQuality
			'ExFor:PdfImageCompression
			'ExFor:PdfCompliance
			'ExSummary:Demonstrates how to save images to PDF using JPEG encoding to decrease file size.
			Dim doc As New Document(MyDir & "SaveOptions.PdfImageComppression.rtf")

			Dim options As New PdfSaveOptions()

			options.ImageCompression = PdfImageCompression.Jpeg
			options.PreserveFormFields = True

			doc.Save(MyDir & "SaveOptions.PdfImageComppression Out.pdf", options)

			Dim optionsA1b As New PdfSaveOptions()
			optionsA1b.Compliance = PdfCompliance.PdfA1b
			optionsA1b.ImageCompression = PdfImageCompression.Jpeg
			optionsA1b.JpegQuality = 100 ' Use JPEG compression at 50% quality to reduce file size.

			doc.Save(MyDir & "SaveOptions.PdfImageComppression PDF_A_1_B Out.pdf", optionsA1b)
			'ExEnd

			Dim optionsA1a As New PdfSaveOptions()
			optionsA1a.Compliance = PdfCompliance.PdfA1a
			optionsA1a.ExportDocumentStructure = True
			optionsA1a.ImageCompression = PdfImageCompression.Jpeg

			doc.Save(MyDir & "SaveOptions.PdfImageComppression PDF_A_1_A Out.pdf", optionsA1a)
			'ExEnd
		End Sub
	End Class
End Namespace
