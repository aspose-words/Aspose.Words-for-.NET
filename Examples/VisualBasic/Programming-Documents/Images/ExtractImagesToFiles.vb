Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Layout
Imports System.Collections
Imports Aspose.Words.Drawing
Public Class ExtractImagesToFiles
    Public Shared Sub Run()
        ' ExStart:ExtractImagesToFiles
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithImages()
        Dim doc As New Document(dataDir & Convert.ToString("Image.SampleImages.doc"))

        Dim shapes As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)
        Dim imageIndex As Integer = 0
        For Each shape As Shape In shapes
            If shape.HasImage Then
                Dim imageFileName As String = String.Format("Image.ExportImages.{0}_out{1}", imageIndex, FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType))
                shape.ImageData.Save(dataDir & imageFileName)
                imageIndex += 1
            End If
        Next
        ' ExEnd:ExtractImagesToFiles
        Console.WriteLine(vbLf & "All images extracted from document.")
    End Sub

End Class
