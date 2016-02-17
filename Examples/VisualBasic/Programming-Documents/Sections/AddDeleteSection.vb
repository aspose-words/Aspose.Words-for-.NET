Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class AddDeleteSection
    Public Shared Sub Run()

        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections() + "Section.AddRemove.doc"
        AddSection(dataDir)
        'DeleteSection(dataDir);
        'DeleteAllSections(dataDir);
    End Sub
    Private Shared Sub AddSection(dataDir As String)
        ' ExStart:AddSection
        Dim doc As New Document(dataDir)
        Dim sectionToAdd As New Section(doc)
        doc.Sections.Add(sectionToAdd)
        ' ExEnd:AddSection
        Console.WriteLine(vbLf & "Section added successfully to the end of the document.")
    End Sub
    Private Shared Sub DeleteSection(dataDir As String)
        ' ExStart:DeleteSection
        Dim doc As New Document(dataDir)
        doc.Sections.RemoveAt(0)
        ' ExEnd:DeleteSection
        Console.WriteLine(vbLf & "Section deleted successfully at 0 index.")
    End Sub
    Private Shared Sub DeleteAllSections(dataDir As String)
        ' ExStart:DeleteAllSections
        Dim doc As New Document(dataDir)
        doc.Sections.Clear()
        ' ExEnd:DeleteAllSections
        Console.WriteLine(vbLf & "All sections deleted successfully form the document.")
    End Sub
End Class
