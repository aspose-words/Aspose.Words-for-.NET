// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
SizeF shapeSizeInDocument = shape.GetShapeRenderer().SizeInPoints;
float width = shapeSizeInDocument.Width; // The width of the shape.
float height = shapeSizeInDocument.Height; // The height of the shape.
            
Size shapeRenderedSize = shape.GetShapeRenderer().GetSizeInPixels(1.0f, 96.0f);

using (Bitmap image = new Bitmap(shapeRenderedSize.Width, shapeRenderedSize.Height))
{
    using (Graphics g = Graphics.FromImage(image))
    {
        // Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.
    }
}
