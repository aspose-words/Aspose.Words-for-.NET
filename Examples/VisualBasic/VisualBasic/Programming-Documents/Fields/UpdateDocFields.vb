Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class UpdateDocFields
    Public Shared Sub Run()
        ' ExStart:UpdateDocFields
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))

        ' This updates all fields in the document.
        doc.UpdateFields()
        dataDir = dataDir & Convert.ToString("Rendering.UpdateFields_out.pdf")
        doc.Save(dataDir)
        ' ExEnd:UpdateDocFields
        Console.WriteLine(Convert.ToString(vbLf & "Fields updated successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
