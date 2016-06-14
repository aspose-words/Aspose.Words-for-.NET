Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Drawing
Imports System.Drawing.Imaging

Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Rendering
Imports Aspose.Words.Saving
Imports Aspose.Words.Tables

Public Class RenderShape
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        ' Load the documents which store the shapes we want to render.
        Dim doc As New Document(dataDir + "TestFile RenderShape.doc")

        ' Retrieve the target shape from the document. In our sample document this is the first shape.
        Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
        
        ' Test rendering of different types of nodes.
        RenderShapeToDisk(dataDir, shape)
        RenderShapeToStream(dataDir, shape)
        RenderShapeToGraphics(dataDir, shape)
        RenderCellToImage(dataDir, doc)
        RenderRowToImage(dataDir, doc)
        RenderParagraphToImage(dataDir, doc)
        FindShapeSizes(shape)
        RenderShapeImage(dataDir, shape)
    End Sub

    Public Shared Sub RenderShapeToDisk(ByVal dataDir As String, ByVal shape As Shape)
        ' ExStart:RenderShapeToDisk
        Dim r As ShapeRenderer = shape.GetShapeRenderer()

        ' Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
        Dim imageOptions As New ImageSaveOptions(SaveFormat.Emf) With {.Scale = 1.5F}

        dataDir = dataDir & "TestFile.RenderToDisk_out_.emf"
        ' Save the rendered image to disk.
        r.Save(dataDir, imageOptions)
        ' ExEnd:RenderShapeToDisk
        Console.WriteLine(vbNewLine & "Shape rendered to disk successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

    Public Shared Sub RenderShapeToStream(ByVal dataDir As String, ByVal shape As Shape)
        ' ExStart:RenderShapeToStream
        Dim r As New ShapeRenderer(shape)

        ' Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
        ' Output the image in gray scale
        ' Reduce the brightness a bit (default is 0.5f).
        Dim imageOptions As New ImageSaveOptions(SaveFormat.Jpeg) With {.ImageColorMode = ImageColorMode.Grayscale, .ImageBrightness = 0.45F}

        dataDir = dataDir & "TestFile.RenderToStream_out_.jpg"
        Dim stream As New FileStream(dataDir, FileMode.Create)

        ' Save the rendered image to the stream using different options.
        r.Save(stream, imageOptions)
        ' ExEnd:RenderShapeToStream
        Console.WriteLine(vbNewLine & "Shape rendered to stream successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

    
    Public Shared Sub RenderShapeToGraphics(ByVal dataDir As String, ByVal shape As Shape)
        ' ExStart:RenderShapeToGraphics
        Dim r As ShapeRenderer = shape.GetShapeRenderer()

        ' Find the size that the shape will be rendered to at the specified scale and resolution.
        Dim shapeSizeInPixels As Size = r.GetSizeInPixels(1.0F, 96.0F)

        ' Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
        ' and make sure that the graphics canvas is large enough to compensate for this.
        Dim maxSide As Integer = System.Math.Max(shapeSizeInPixels.Width, shapeSizeInPixels.Height)

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
        ' ExEnd:RenderShapeToGraphics
    End Sub

    Public Shared Sub RenderCellToImage(ByVal dataDir As String, ByVal doc As Document)
        ' ExStart:RenderCellToImage
        Dim cell As Cell = CType(doc.GetChild(NodeType.Cell, 2, True), Cell) ' The third cell in the first table.
        dataDir = dataDir & "TestFile.RenderCell_out_.png"
        RenderNode(cell, dataDir, Nothing)
        ' ExEnd:RenderCellToImage
        Console.WriteLine(vbNewLine & "Cell rendered to image successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

    Public Shared Sub RenderRowToImage(ByVal dataDir As String, ByVal doc As Document)
        ' ExStart:RenderRowToImage
        Dim row As Row = CType(doc.GetChild(NodeType.Row, 0, True), Row) ' The first row in the first table.
        dataDir = dataDir & "TestFile.RenderRow_out_.png"
        RenderNode(row, dataDir, Nothing)
        ' ExEnd:RenderRowToImage
        Console.WriteLine(vbNewLine & "Row rendered to image successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

    Public Shared Sub RenderParagraphToImage(ByVal dataDir As String, ByVal doc As Document)
        ' ExStart:RenderParagraphToImage
        Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
        Dim paragraph As Paragraph = CType(shape.LastParagraph, Paragraph)

        ' Save the node with a light pink background.
        Dim options As New ImageSaveOptions(SaveFormat.Png)
        options.PaperColor = Color.LightPink
        dataDir = dataDir & "TestFile.RenderParagraph_out_.png"
        RenderNode(paragraph, dataDir, options)
        ' ExEnd:RenderParagraphToImage
        Console.WriteLine(vbNewLine & "Paragraph rendered to image successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

    Public Shared Sub FindShapeSizes(ByVal shape As Shape)
        ' ExStart:FindShapeSizes
        Dim shapeSizeInDocument As SizeF = shape.GetShapeRenderer().SizeInPoints
        Dim width As Single = shapeSizeInDocument.Width ' The width of the shape.
        Dim height As Single = shapeSizeInDocument.Height ' The height of the shape.
        
        Dim shapeRenderedSize As Size = shape.GetShapeRenderer().GetSizeInPixels(1.0F, 96.0F)

        Using image As New Bitmap(shapeRenderedSize.Width, shapeRenderedSize.Height)
            Using g As Graphics = Graphics.FromImage(image)
                ' Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.
            End Using
        End Using
        ' ExEnd:FindShapeSizes
    End Sub
    Public Shared Sub RenderShapeImage(dataDir As String, shape As Shape)
        ' ExStart:RenderShapeImage
        dataDir = dataDir & Convert.ToString("TestFile.RenderShape_out_.jpg")
        ' Save the Shape image to disk in JPEG format and using default options.
        shape.GetShapeRenderer().Save(dataDir, Nothing)
        ' ExEnd:RenderShapeImage
        Console.WriteLine(Convert.ToString(vbLf & "Shape image rendered successfully." & vbLf & "File saved at ") & dataDir)
    End Sub

    ''' <summary>
    ''' Renders any node in a document to the path specified using the image save options.
    ''' </summary>
    ''' <param name="node">The node to render.</param>
    ''' <param name="filepath">The path to save the rendered image to.</param>
    ''' <param name="imageOptions">The image options to use during rendering. This can be null.</param>
    Public Shared Sub RenderNode(ByVal node As Node, ByVal filePath As String, ByVal imageOptions As ImageSaveOptions)
        ' This code is taken from public API samples of AW.
        ' Previously to find opaque bounds of the shape the function
        ' that checks every pixel of the rendered image was used.
        ' For now opaque bounds is got using ShapeRenderer.GetOpaqueRectangleInPixels method.

        ' If no image options are supplied, create default options.
        If imageOptions Is Nothing Then
            imageOptions = New ImageSaveOptions(FileFormatUtil.ExtensionToSaveFormat(Path.GetExtension(filePath)))
        End If

        ' Store the paper color to be used on the final image and change to transparent.
        ' This will cause any content around the rendered node to be removed later on.
        Dim savePaperColor As Color = imageOptions.PaperColor
        imageOptions.PaperColor = Color.Transparent

        ' There a bug which affects the cache of a cloned node. To avoid this we instead clone the entire document including all nodes,
        ' find the matching node in the cloned document and render that instead.
        Dim doc As Document = CType(node.Document.Clone(True), Document)
        node = doc.GetChild(NodeType.Any, node.Document.GetChildNodes(NodeType.Any, True).IndexOf(node), True)

        ' Create a temporary shape to store the target node in. This shape will be rendered to retrieve
        ' the rendered content of the node.
        Dim shape As Shape = New Shape(doc, ShapeType.TextBox)
        Dim parentSection As Section = CType(node.GetAncestor(NodeType.Section), Section)

        ' Assume that the node cannot be larger than the page in size.
        shape.Width = parentSection.PageSetup.PageWidth
        shape.Height = parentSection.PageSetup.PageHeight
        shape.FillColor = Color.Transparent ' We must make the shape and paper color transparent.

        ' Don't draw a surronding line on the shape.
        shape.Stroked = False

        ' Move up through the DOM until we find node which is suitable to insert into a Shape (a node with a parent can contain paragraph, tables the same as a shape).
        ' Each parent node is cloned on the way up so even a descendant node passed to this method can be rendered.
        ' Since we are working with the actual nodes of the document we need to clone the target node into the temporary shape.
        Dim currentNode As Node = node
        Do While Not (TypeOf currentNode.ParentNode Is InlineStory OrElse TypeOf currentNode.ParentNode Is Story OrElse TypeOf currentNode.ParentNode Is ShapeBase)
            Dim parent As CompositeNode = CType(currentNode.ParentNode.Clone(False), CompositeNode)
            currentNode = currentNode.ParentNode
            parent.AppendChild(node.Clone(True))
            node = parent ' Store this new node to be inserted into the shape.
        Loop

        ' We must add the shape to the document tree to have it rendered.
        shape.AppendChild(node.Clone(True))
        parentSection.Body.FirstParagraph.AppendChild(shape)

        ' Render the shape to stream so we can take advantage of the effects of the ImageSaveOptions class.
        ' Retrieve the rendered image and remove the shape from the document.
        Dim stream As MemoryStream = New MemoryStream()
        Dim renderer As ShapeRenderer = shape.GetShapeRenderer()
        renderer.Save(stream, imageOptions)
        shape.Remove()

        Dim crop As Rectangle = renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.Resolution)

        ' Load the image into a new bitmap.
        Using renderedImage As Bitmap = New Bitmap(stream)
            Dim croppedImage As Bitmap = New Bitmap(crop.Width, crop.Height)
            croppedImage.SetResolution(imageOptions.Resolution, imageOptions.Resolution)

            ' Create the final image with the proper background color.
            Using g As Graphics = Graphics.FromImage(croppedImage)
                g.Clear(savePaperColor)
                g.DrawImage(renderedImage, New Rectangle(0, 0, croppedImage.Width, croppedImage.Height), crop.X, crop.Y, crop.Width, crop.Height, GraphicsUnit.Pixel)

                croppedImage.Save(filePath)
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Finds the minimum bounding box around non-transparent pixels in a Bitmap.
    ''' </summary>
    Public Shared Function FindBoundingBoxAroundNode(ByVal originalBitmap As Bitmap) As Rectangle
        Dim min As New Point(Integer.MaxValue, Integer.MaxValue)
        Dim max As New Point(Integer.MinValue, Integer.MinValue)

        For x As Integer = 0 To originalBitmap.Width - 1
            For y As Integer = 0 To originalBitmap.Height - 1
                ' Note that you can speed up this part of the algorithm by using LockBits and unsafe code instead of GetPixel.
                Dim pixelColor As Color = originalBitmap.GetPixel(x, y)

                ' For each pixel that is not transparent calculate the bounding box around it.
                If pixelColor.ToArgb() <> Color.Empty.ToArgb() Then
                    min.X = System.Math.Min(x, min.X)
                    min.Y = System.Math.Min(y, min.Y)
                    max.X = System.Math.Max(x, max.X)
                    max.Y = System.Math.Max(y, max.Y)
                End If
            Next y
        Next x

        ' Add one pixel to the width and height to avoid clipping.
        Return New Rectangle(min.X, min.Y, (max.X - min.X) + 1, (max.Y - min.Y) + 1)
    End Function
End Class
