// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithImages();
Document doc = new Document(dataDir + "Image.SampleImages.doc");

NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
int imageIndex = 0;
foreach (Shape shape in shapes)
{
    if (shape.HasImage)
    {
        string imageFileName = string.Format(
            "Image.ExportImages.{0}_out_{1}", imageIndex, FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType));
        shape.ImageData.Save(dataDir + imageFileName);
        imageIndex++;
    }
}
