' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Public Class HandleMergeImageFieldFromBlob
    Implements IFieldMergingCallback
    Private Sub IFieldMergingCallback_FieldMerging(args As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
        ' Do nothing.
    End Sub

    ''' <summary>
    ''' This is called when mail merge engine encounters Image:XXX merge field in the document.
    ''' You have a chance to return an Image object, file name or a stream that contains the image.
    ''' </summary>
    Private Sub IFieldMergingCallback_ImageFieldMerging(e As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
        ' The field value is a byte array, just cast it and create a stream on it.
        Dim imageStream As New MemoryStream(DirectCast(e.FieldValue, Byte()))
        ' Now the mail merge engine will retrieve the image from the stream.
        e.ImageStream = imageStream
    End Sub
End Class
