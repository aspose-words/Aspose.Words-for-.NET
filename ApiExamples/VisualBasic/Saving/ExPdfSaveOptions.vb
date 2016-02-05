Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports Aspose.Words.Saving
Imports NUnit.Framework


Namespace ApiExamples.Saving
	<TestFixture> _
	Friend Class ExPdfSaveOptions
		Inherits ApiExampleBase
		<Test> _
		Public Sub CreateMissingOutlineLevels()
			'ExStart
			'ExFor:Saving.PdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels
			'ExSummary:Shows how to create missing outline levels saving the document in pdf
			Dim doc As New Aspose.Words.Document()

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

			doc.Save(MyDir & "CreateMissingOutlineLevels.pdf", pdfSaveOptions)
			'ExEnd
		End Sub
	End Class
End Namespace
