' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim r As ShapeRenderer = shape.GetShapeRenderer()

' Find the size that the shape will be rendered to at the specified scale and resolution.
Dim shapeSizeInPixels As Size = r.GetSizeInPixels(1.0F, 96.0F)

' Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
' and make sure that the graphics canvas is large enough to compensate for this.
Dim maxSide As Integer = Math.Max(shapeSizeInPixels.Width, shapeSizeInPixels.Height)

Using image As New Bitmap(CInt(Fix(maxSide * 1.25)), CInt(Fix(maxSide * 1.25)))
    ' Rendering to a graphics object means we can specify settings and transformations to be applied to 
    ' the shape that is rendered. In our case we will rotate the rendered shape.
    Using gr As Graphics = Graphics.FromImage(image)
        ' Clear the shape with the background color of the document.
        gr.Clear(Color.White)
        ' Center the rotation using translation method below
        gr.TranslateTransform(CSng(image.Width) / 8, CSng(image.Height) / 2)
        ' Rotate the image by 45 degrees.
        gr.RotateTransform(45)
        ' Undo the translation.
        gr.TranslateTransform(-CSng(image.Width) / 8, -CSng(image.Height) / 2)

        ' Render the shape onto the graphics object.
        r.RenderToSize(gr, 0, 0, shapeSizeInPixels.Width, shapeSizeInPixels.Height)
    End Using
    dataDir = dataDir & "TestFile.RenderToGraphics_out_.png"
    image.Save(dataDir, ImageFormat.Png)
    Console.WriteLine(vbNewLine & "Shape rendered to graphics successfully." & vbNewLine & "File saved at " + dataDir)
End Using
