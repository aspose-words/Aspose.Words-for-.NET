Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class AppendSectionContent
    Public Shared Sub Run()
        ' ExStart:AppendSectionContent
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections()
        Dim doc As New Document(dataDir & Convert.ToString("Section.AppendContent.doc"))
        ' This is the section that we will append and prepend to.
        Dim section As Section = doc.Sections(2)

        ' This copies content of the 1st section and inserts it at the beginning of the specified section.
        Dim sectionToPrepend As Section = doc.Sections(0)
        section.PrependContent(sectionToPrepend)

        ' This copies content of the 2nd section and inserts it at the end of the specified section.
        Dim sectionToAppend As Section = doc.Sections(1)
        section.AppendContent(sectionToAppend)
        ' ExEnd:AppendSectionContent
        Console.WriteLine(vbLf & "Section content appended successfully.")
    End Sub
End Class
