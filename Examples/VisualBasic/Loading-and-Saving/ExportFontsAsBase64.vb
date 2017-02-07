Imports Aspose.Words.Saving
Public Class ExportFontsAsBase64
    Public Shared Sub Run()
        ' ExStart:ExportFontsAsBase64            
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()
        Dim fileName As String = "Document.doc"
        Dim doc As New Document(dataDir & fileName)
        Dim saveOptions As New HtmlSaveOptions()
        saveOptions.ExportFontResources = True
        saveOptions.ExportFontsAsBase64 = True
        dataDir = dataDir & Convert.ToString("ExportFontsAsBase64_out.html")
        doc.Save(dataDir, saveOptions)
        ' ExEnd:ExportFontsAsBase64
        Console.WriteLine(Convert.ToString(vbLf & "Save option specified successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class