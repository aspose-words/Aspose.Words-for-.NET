' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim r As ShapeRenderer = shape.GetShapeRenderer()

' Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
Dim imageOptions As New ImageSaveOptions(SaveFormat.Emf) With {.Scale = 1.5F}

dataDir = dataDir & "TestFile.RenderToDisk_out_.emf"
' Save the rendered image to disk.
r.Save(dataDir, imageOptions)
