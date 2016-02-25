' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithImages()
Dim doc As New Document(dataDir & Convert.ToString("Image.SampleImages.doc"))

Dim shapes As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)
Dim imageIndex As Integer = 0
For Each shape As Shape In shapes
    If shape.HasImage Then
        Dim imageFileName As String = String.Format("Image.ExportImages.{0}_out_{1}", imageIndex, FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType))
        shape.ImageData.Save(dataDir & imageFileName)
        imageIndex += 1
    End If
Next
