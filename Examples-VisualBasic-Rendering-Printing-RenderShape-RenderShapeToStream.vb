' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim r As New ShapeRenderer(shape)

' Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
' Output the image in gray scale
' Reduce the brightness a bit (default is 0.5f).
Dim imageOptions As New ImageSaveOptions(SaveFormat.Jpeg) With {.ImageColorMode = ImageColorMode.Grayscale, .ImageBrightness = 0.45F}

dataDir = dataDir & "TestFile.RenderToStream_out_.jpg"
Dim stream As New FileStream(dataDir, FileMode.Create)

' Save the rendered image to the stream using different options.
r.Save(stream, imageOptions)
