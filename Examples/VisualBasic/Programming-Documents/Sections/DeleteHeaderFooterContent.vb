Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class DeleteHeaderFooterContent
    Public Shared Sub Run()
        ' ExStart:DeleteHeaderFooterContent
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections()

        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
        Dim section As Section = doc.Sections(0)
        section.ClearHeadersFooters()
        ' ExEnd:DeleteHeaderFooterContent
        Console.WriteLine(vbLf & "Header and footer content of 0 index deleted successfully.")
    End Sub
End Class
