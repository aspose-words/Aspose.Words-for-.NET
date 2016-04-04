' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim shapeSizeInDocument As SizeF = shape.GetShapeRenderer().SizeInPoints
Dim width As Single = shapeSizeInDocument.Width ' The width of the shape.
Dim height As Single = shapeSizeInDocument.Height ' The height of the shape.
        
Dim shapeRenderedSize As Size = shape.GetShapeRenderer().GetSizeInPixels(1.0F, 96.0F)

Using image As New Bitmap(shapeRenderedSize.Width, shapeRenderedSize.Height)
    Using g As Graphics = Graphics.FromImage(image)
        ' Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.
    End Using
End Using
