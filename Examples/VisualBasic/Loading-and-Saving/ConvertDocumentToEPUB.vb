Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Saving

Public Class ConvertDocumentToEPUB
    Public Shared Sub Run()
        ' ExStart:ConvertDocumentToEPUB
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' Load the document from disk.
        Dim doc As New Document(dataDir & Convert.ToString("Test File (doc).doc"))

        ' Create a new instance of HtmlSaveOptions. This object allows us to set options that control
        ' how the output document is saved.
        Dim saveOptions As New HtmlSaveOptions()

        ' Specify the desired encoding.
        saveOptions.Encoding = System.Text.Encoding.UTF8

        ' Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
        ' which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
        ' HTML files greater than a certain size e.g 300kb.
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph

        ' Specify that we want to export document properties.
        saveOptions.ExportDocumentProperties = True

        ' Specify that we want to save in EPUB format.
        saveOptions.SaveFormat = SaveFormat.Epub

        ' Export the document as an EPUB file.
        doc.Save(dataDir & Convert.ToString("Document.EpubConversion_out_.epub"), saveOptions)

        ' ExEnd:ConvertDocumentToEPUB

        Console.WriteLine(vbLf & "Document converted to EPUB successfully.")
    End Sub
End Class
