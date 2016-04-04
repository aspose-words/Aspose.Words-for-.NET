// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
ShapeRenderer r = new ShapeRenderer(shape);

// Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
{
    // Output the image in gray scale
    ImageColorMode = ImageColorMode.Grayscale,

    // Reduce the brightness a bit (default is 0.5f).
    ImageBrightness = 0.45f
};
dataDir = dataDir + "TestFile.RenderToStream_out_.jpg";
FileStream stream = new FileStream(dataDir, FileMode.Create);

// Save the rendered image to the stream using different options.
r.Save(stream, imageOptions);
