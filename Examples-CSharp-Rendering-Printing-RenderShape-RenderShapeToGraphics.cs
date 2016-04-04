// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
ShapeRenderer r = shape.GetShapeRenderer();

// Find the size that the shape will be rendered to at the specified scale and resolution.
Size shapeSizeInPixels = r.GetSizeInPixels(1.0f, 96.0f);

// Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
// and make sure that the graphics canvas is large enough to compensate for this.
int maxSide = Math.Max(shapeSizeInPixels.Width, shapeSizeInPixels.Height);

using (Bitmap image = new Bitmap((int)(maxSide * 1.25), (int)(maxSide * 1.25)))
{
    // Rendering to a graphics object means we can specify settings and transformations to be applied to 
    // the shape that is rendered. In our case we will rotate the rendered shape.
    using (Graphics gr = Graphics.FromImage(image))
    {
        // Clear the shape with the background color of the document.
        gr.Clear(shape.Document.PageColor);
        // Center the rotation using translation method below
        gr.TranslateTransform((float)image.Width / 8, (float)image.Height / 2);
        // Rotate the image by 45 degrees.
        gr.RotateTransform(45);
        // Undo the translation.
        gr.TranslateTransform(-(float)image.Width / 8, -(float)image.Height / 2);

        // Render the shape onto the graphics object.
        r.RenderToSize(gr, 0, 0, shapeSizeInPixels.Width, shapeSizeInPixels.Height);
    }
    dataDir = dataDir + "TestFile.RenderToGraphics_out_.png";
    image.Save(dataDir, ImageFormat.Png);
    Console.WriteLine("\nShape rendered to graphics successfully.\nFile saved at " + dataDir);
}
