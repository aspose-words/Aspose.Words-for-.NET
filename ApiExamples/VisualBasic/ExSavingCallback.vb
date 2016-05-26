Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Saving
Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Friend Class ExSavingCallback
		Inherits ApiExampleBase
		<Test> _
		Public Sub CheckThatAllMethodsArePresent()
			Dim htmlFixedSaveOptions As New HtmlFixedSaveOptions()
			htmlFixedSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()

			Dim imageSaveOptions As New ImageSaveOptions(SaveFormat.Png)
			imageSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()

			Dim pdfSaveOptions As New PdfSaveOptions()
			pdfSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()

			Dim psSaveOptions As New PsSaveOptions()
			psSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()

			Dim svgSaveOptions As New SvgSaveOptions()
			svgSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()

			Dim swfSaveOptions As New SwfSaveOptions()
			swfSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()

			Dim xamlFixedSaveOptions As New XamlFixedSaveOptions()
			xamlFixedSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()

			Dim xpsSaveOptions As New XpsSaveOptions()
			xpsSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()
		End Sub

		<Test> _
		Public Sub PageFileNameSavingCallback()
			Dim doc As New Document(MyDir & "Rendering.doc")

			Dim htmlFixedSaveOptions As HtmlFixedSaveOptions = New HtmlFixedSaveOptions With {.PageIndex = 0, .PageCount = doc.PageCount}
			htmlFixedSaveOptions.PageSavingCallback = New CustomPageFileNamePageSavingCallback()

			doc.Save(MyDir & "\Artifacts\out.html", htmlFixedSaveOptions)

			Dim filePaths() As String = Directory.GetFiles(MyDir, "Page_*.html")

			For i As Integer = 0 To doc.PageCount - 1
				Dim file As String = String.Format(MyDir & "Page_{0}.html", i)
				Assert.AreEqual(file, filePaths(i))
			Next i
		End Sub

		<Test> _
		Public Sub PageStreamSavingCallback()
			Dim docStream As Stream = New FileStream(MyDir & "Rendering.doc", FileMode.Open)
			Dim doc As New Document(docStream)

			Dim htmlFixedSaveOptions As HtmlFixedSaveOptions = New HtmlFixedSaveOptions With {.PageIndex = 0, .PageCount = doc.PageCount}
			htmlFixedSaveOptions.PageSavingCallback = New CustomPageStreamPageSavingCallback()

			doc.Save(MyDir & "\Artifacts\out.html", htmlFixedSaveOptions)

			docStream.Close()
		End Sub

		''' <summary>
		''' Custom PageFileName is specified.
		''' </summary>
		Private Class CustomPageFileNamePageSavingCallback
			Implements IPageSavingCallback
            Public Sub PageSaving(ByVal args As PageSavingArgs) Implements IPageSavingCallback.PageSaving
                ' Specify name of the output file for the current page.
                args.PageFileName = String.Format(MyDir & "Page_{0}.html", args.PageIndex)
            End Sub
		End Class

		''' <summary>
		''' Custom PageStream is specified.
		''' </summary>
		Private Class CustomPageStreamPageSavingCallback
			Implements IPageSavingCallback
            Public Sub PageSaving(ByVal args As PageSavingArgs) Implements IPageSavingCallback.PageSaving
                ' Specify memory stream for the current page.
                args.PageStream = New MemoryStream()
                args.KeepPageStreamOpen = True
            End Sub
		End Class
	End Class
End Namespace
