Imports Microsoft.VisualBasic
Imports System.Text
Imports Aspose.Words
Imports Aspose.Words.Saving
Imports NUnit.Framework


Namespace ApiExamples.Saving
	<TestFixture> _
	Friend Class ExHtmlFixedSaveOptions
		Inherits ApiExampleBase
		<Test> _
		Public Sub UseEncoding()
			'ExStart
			'ExFor:Saving.HtmlFixedSaveOptions.Encoding
			'ExSummary:Shows how to use "Encoding" parameter with "HtmlFixedSaveOptions"
			Dim doc As New Aspose.Words.Document()

			Dim builder As New DocumentBuilder(doc)
			builder.Writeln("Hello World!")

			'Create "HtmlFixedSaveOptions" with "Encoding" parameter
			'You can also set "Encoding" using System.Text.Encoding, like "Encoding.ASCII", or "Encoding.GetEncoding()"
			Dim htmlFixedSaveOptions As HtmlFixedSaveOptions = New HtmlFixedSaveOptions With {.Encoding = New ASCIIEncoding(), .SaveFormat = SaveFormat.HtmlFixed}

			'Uses "HtmlFixedSaveOptions"
			doc.Save(MyDir & "UseEncoding.html", htmlFixedSaveOptions)
			'ExEnd
		End Sub
	End Class
End Namespace
