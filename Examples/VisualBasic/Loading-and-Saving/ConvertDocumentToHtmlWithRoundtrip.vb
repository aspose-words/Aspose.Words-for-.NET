Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class ConvertDocumentToHtmlWithRoundtrip
    Public Shared Sub Run()
        ' ExStart:ConvertDocumentToHtmlWithRoundtrip
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' Load the document from disk.
        Dim doc As New Document(dataDir & Convert.ToString("Test File (doc).doc"))

        Dim options As New HtmlSaveOptions()

        'HtmlSaveOptions.ExportRoundtripInformation property specifies
        'whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
        'Default value is true for HTML and false for MHTML and EPUB.
        options.ExportRoundtripInformation = True

        doc.Save(dataDir & Convert.ToString("ExportRoundtripInformation_out_.html"), options)
        ' ExEnd:ConvertDocumentToHtmlWithRoundtrip

        Console.WriteLine(vbLf & "Document converted to html with roundtrip informations successfully.")
    End Sub
End Class
