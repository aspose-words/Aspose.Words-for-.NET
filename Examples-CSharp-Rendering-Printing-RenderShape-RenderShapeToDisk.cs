// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
ShapeRenderer r = shape.GetShapeRenderer();

// Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Emf)
{
    Scale = 1.5f
};

dataDir = dataDir + "TestFile.RenderToDisk_out_.emf";
// Save the rendered image to disk.
r.Save(dataDir, imageOptions);
