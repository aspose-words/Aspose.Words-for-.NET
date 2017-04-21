Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class DeleteSectionContent
    Public Shared Sub Run()
        ' ExStart:DeleteSectionContent
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections()

        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
        Dim section As Section = doc.Sections(0)
        section.ClearContent()
        ' ExEnd:DeleteSectionContent
        Console.WriteLine(vbLf & "Section content at 0 index deleted successfully.")
    End Sub
End Class
