Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class CopySection
    Public Shared Sub Run()
        ' ExStart:CopySection
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections()

        Dim srcDoc As New Document(dataDir & Convert.ToString("Document.doc"))
        Dim dstDoc As New Document()

        Dim sourceSection As Section = srcDoc.Sections(0)
        Dim newSection As Section = DirectCast(dstDoc.ImportNode(sourceSection, True), Section)
        dstDoc.Sections.Add(newSection)
        dataDir = dataDir & Convert.ToString("Document.Copy_out_.doc")
        dstDoc.Save(dataDir)
        ' ExEnd:CopySection
        Console.WriteLine(Convert.ToString(vbLf & "Section copied successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
