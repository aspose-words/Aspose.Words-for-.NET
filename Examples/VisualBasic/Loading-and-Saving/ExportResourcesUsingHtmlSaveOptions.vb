Imports Aspose.Words.Saving
Public Class ExportResourcesUsingHtmlSaveOptions
    Public Shared Sub Run()
        ' ExStart:ExportResourcesUsingHtmlSaveOptions            
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()
        Dim fileName As String = "Document.doc"
        Dim doc As New Document(dataDir & fileName)
        Dim saveOptions As New HtmlSaveOptions()
        saveOptions.CssStyleSheetType = CssStyleSheetType.External
        saveOptions.ExportFontResources = True
        saveOptions.ResourceFolder = dataDir & Convert.ToString("\Resources")
        saveOptions.ResourceFolderAlias = "http://example.com/resources"
        doc.Save(dataDir & Convert.ToString("ExportResourcesUsingHtmlSaveOptions.html"), saveOptions)
        ' ExEnd:ExportResourcesUsingHtmlSaveOptions
        Console.WriteLine(Convert.ToString(vbLf & "Save option specified successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class