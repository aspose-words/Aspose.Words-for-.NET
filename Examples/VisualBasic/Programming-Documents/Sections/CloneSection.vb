Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class CloneSection
    Public Shared Sub Run()
        ' ExStart:CloneSection
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections()

        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
        Dim cloneSection As Section = doc.Sections(0).Clone()
        ' ExEnd:CloneSection
        Console.WriteLine(vbLf & "0 index section clone successfully.")
    End Sub
End Class
