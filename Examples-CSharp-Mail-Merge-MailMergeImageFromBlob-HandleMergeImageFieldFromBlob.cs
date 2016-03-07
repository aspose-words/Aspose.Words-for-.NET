// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
public class HandleMergeImageFieldFromBlob : IFieldMergingCallback
{
    void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
    {
        // Do nothing.
    }

    /// <summary>
    /// This is called when mail merge engine encounters Image:XXX merge field in the document.
    /// You have a chance to return an Image object, file name or a stream that contains the image.
    /// </summary>
    void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
    {
        // The field value is a byte array, just cast it and create a stream on it.
        MemoryStream imageStream = new MemoryStream((byte[])e.FieldValue);
        // Now the mail merge engine will retrieve the image from the stream.
        e.ImageStream = imageStream;
    }
}
