Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words

Public Class LoadAndSaveToDisk
    Public Shared Sub Run()
        ' ExStart:LoadAndSave
        ' ExStart:OpenDocument
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()
        Dim fileName As String = "Document.doc"
        ' Load the document from the absolute path on disk.
        Dim doc As New Document(dataDir & fileName)
        ' ExEnd:OpenDocument

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the document as DOCX document.");
        doc.Save(dataDir)
        ' ExEnd:LoadAndSave
        Console.WriteLine(vbNewLine + "Existing document loaded and saved successfully." + vbNewLine + "File saved at " + dataDir)
    End Sub
End Class
