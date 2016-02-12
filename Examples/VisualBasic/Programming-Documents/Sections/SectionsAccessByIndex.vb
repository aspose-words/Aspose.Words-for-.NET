Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class SectionsAccessByIndex
    Public Shared Sub Run()
        ' ExStart:SectionsAccessByIndex
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections()
        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
        Dim section As Section = doc.Sections(0)
        section.PageSetup.LeftMargin = 90
        ' 3.17 cm
        section.PageSetup.RightMargin = 90
        ' 3.17 cm
        section.PageSetup.TopMargin = 72
        ' 2.54 cm
        section.PageSetup.BottomMargin = 72
        ' 2.54 cm
        section.PageSetup.HeaderDistance = 35.4
        ' 1.25 cm
        section.PageSetup.FooterDistance = 35.4
        ' 1.25 cm
        section.PageSetup.TextColumns.Spacing = 35.4
        ' 1.25 cm
        ' ExEnd:SectionsAccessByIndex
        Console.WriteLine(vbLf & "Section at 0 index have text " + "'" + section.GetText() + "'")
    End Sub
End Class
