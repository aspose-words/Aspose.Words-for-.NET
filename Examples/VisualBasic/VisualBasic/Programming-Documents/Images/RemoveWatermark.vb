Imports Aspose.Words.Drawing

Public Class RemoveWatermark
    ' ExStart:RemoveWatermark
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithImages()
        Dim fileName As String = "RemoveWatermark.docx"
        Dim doc As New Document(dataDir & fileName)
        RemoveWatermarkText(doc)
        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        doc.Save(dataDir)
    End Sub

    Private Shared Sub RemoveWatermarkText(doc As Document)
        For Each hf As HeaderFooter In doc.GetChildNodes(NodeType.HeaderFooter, True)
            For Each shape As Shape In hf.GetChildNodes(NodeType.Shape, True)
                If shape.Name.Contains("WaterMark") Then
                    shape.Remove()
                End If
            Next
        Next
    End Sub
    ' ExEnd:RemoveWatermark
End Class
