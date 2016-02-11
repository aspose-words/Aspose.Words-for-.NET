Imports Microsoft.VisualBasic
Imports Aspose.Words
Imports Aspose.Words.Settings
Public Class SetViewOption
   Public Shared Sub Run()
        ' ExStart:SetViewOption
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Load the template document.
        Dim doc As New Document(dataDir & Convert.ToString("TestFile.doc"))
        ' Set view option.
        doc.ViewOptions.ViewType = ViewType.PageLayout
        doc.ViewOptions.ZoomPercent = 50

        dataDir = dataDir & Convert.ToString("TestFile.SetZoom_out_.doc")
        ' Save the finished document.
        doc.Save(dataDir)
        ' ExEnd:SetViewOption

        Console.WriteLine(Convert.ToString(vbLf & "View option setup successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
