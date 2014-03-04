' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

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

Namespace RenderShapes
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Load the documents which store the shapes we want to render.
			Dim doc As New Document(dataDir & "TestFile.doc")
			Dim doc2 As New Document(dataDir & "TestFile.docx")

			' Retrieve the target shape from the document. In our sample document this is the first shape.
			Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
			Dim drawingML As DrawingML = CType(doc2.GetChild(NodeType.DrawingML, 0, True), DrawingML)

			' Test rendering of different types of nodes.
			RenderShapeToDisk(dataDir, shape)
			RenderShapeToStream(dataDir, shape)
			RenderShapeToGraphics(dataDir, shape)
			RenderDrawingMLToDisk(dataDir, drawingML)
			RenderCellToImage(dataDir, doc)
			RenderRowToImage(dataDir, doc)
			RenderParagraphToImage(dataDir, doc)
			FindShapeSizes(shape)
		End Sub

		Public Shared Sub RenderShapeToDisk(ByVal dataDir As String, ByVal shape As Shape)
			'ExStart
			'ExFor:ShapeRenderer
			'ExFor:ShapeBase.GetShapeRenderer
			'ExFor:ImageSaveOptions
			'ExFor:ImageSaveOptions.Scale
			'ExFor:ShapeRenderer.Save(String, ImageSaveOptions)
			'ExId:RenderShapeToDisk
			'ExSummary:Shows how to render a shape independent of the document to an EMF image and save it to disk.
			' The shape render is retrieved using this method. This is made into a separate object from the shape as it internally
			' caches the rendered shape.
			Dim r As ShapeRenderer = shape.GetShapeRenderer()

			' Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
			Dim imageOptions As New ImageSaveOptions(SaveFormat.Emf) With {.Scale = 1.5f}

			' Save the rendered image to disk.
			r.Save(dataDir & "TestFile.RenderToDisk Out.emf", imageOptions)
			'ExEnd
		End Sub

		Public Shared Sub RenderShapeToStream(ByVal dataDir As String, ByVal shape As Shape)
			'ExStart
			'ExFor:ShapeRenderer
			'ExFor:ShapeRenderer.#ctor(ShapeBase)
			'ExFor:ImageSaveOptions.ImageColorMode
			'ExFor:ImageSaveOptions.ImageBrightness
			'ExFor:ShapeRenderer.Save(Stream, ImageSaveOptions)
			'ExId:RenderShapeToStream
			'ExSummary:Shows how to render a shape independent of the document to a JPEG image and save it to a stream.
			' We can also retrieve the renderer for a shape by using the ShapeRenderer constructor.
			Dim r As New ShapeRenderer(shape)

			' Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
				' Output the image in gray scale
				' Reduce the brightness a bit (default is 0.5f).
			Dim imageOptions As New ImageSaveOptions(SaveFormat.Jpeg) With {.ImageColorMode = ImageColorMode.Grayscale, .ImageBrightness = 0.45f}

			Dim stream As New FileStream(dataDir & "TestFile.RenderToStream Out.jpg", FileMode.CreateNew)

			' Save the rendered image to the stream using different options.
			r.Save(stream, imageOptions)
			'ExEnd
		End Sub

		Public Shared Sub RenderDrawingMLToDisk(ByVal dataDir As String, ByVal drawingML As DrawingML)
			'ExStart
			'ExFor:DrawingML.GetShapeRenderer
			'ExFor:ShapeRenderer.Save(String, ImageSaveOptions)
			'ExFor:DrawingML
			'ExId:RenderDrawingMLToDisk
			'ExSummary:Shows how to render a DrawingML image independent of the document to a JPEG image on the disk.
			' Save the DrawingML image to disk in JPEG format and using default options.
			drawingML.GetShapeRenderer().Save(dataDir & "TestFile.RenderDrawingML Out.jpg", Nothing)
			'ExEnd
		End Sub

		Public Shared Sub RenderShapeToGraphics(ByVal dataDir As String, ByVal shape As Shape)
			'ExStart
			'ExFor:ShapeRenderer
			'ExFor:ShapeBase.GetShapeRenderer
			'ExFor:ShapeRenderer.GetSizeInPixels
			'ExFor:ShapeRenderer.RenderToSize
			'ExId:RenderShapeToGraphics
			'ExSummary:Shows how to render a shape independent of the document to a .NET Graphics object and apply rotation to the rendered image.
			' The shape renderer is retrieved using this method. This is made into a separate object from the shape as it internally
			' caches the rendered shape.
			Dim r As ShapeRenderer = shape.GetShapeRenderer()

			' Find the size that the shape will be rendered to at the specified scale and resolution.
			Dim shapeSizeInPixels As Size = r.GetSizeInPixels(1.0f, 96.0f)

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

				image.Save(dataDir & "TestFile.RenderToGraphics.png", ImageFormat.Png)
			End Using
			'ExEnd
		End Sub

		Public Shared Sub RenderCellToImage(ByVal dataDir As String, ByVal doc As Document)
			'ExStart
			'ExId:RenderCellToImage
			'ExSummary:Shows how to render a cell of a table independent of the document.
			Dim cell As Cell = CType(doc.GetChild(NodeType.Cell, 2, True), Cell) ' The third cell in the first table.
			RenderNode(cell, dataDir & "TestFile.RenderCell Out.png", Nothing)
			'ExEnd
		End Sub

		Public Shared Sub RenderRowToImage(ByVal dataDir As String, ByVal doc As Document)
			'ExStart
			'ExId:RenderRowToImage
			'ExSummary:Shows how to render a row of a table independent of the document.
			Dim row As Row = CType(doc.GetChild(NodeType.Row, 0, True), Row) ' The first row in the first table.
			RenderNode(row, dataDir & "TestFile.RenderRow Out.png", Nothing)
			'ExEnd
		End Sub

		Public Shared Sub RenderParagraphToImage(ByVal dataDir As String, ByVal doc As Document)
			'ExStart
			'ExFor:Shape.LastParagraph
			'ExId:RenderParagraphToImage
			'ExSummary:Shows how to render a paragraph with a custom background color independent of the document. 
			' Retrieve the first paragraph in the main shape.
			Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
			Dim paragraph As Paragraph = CType(shape.LastParagraph, Paragraph)

			' Save the node with a light pink background.
			Dim options As New ImageSaveOptions(SaveFormat.Png)
			options.PaperColor = Color.LightPink

			RenderNode(paragraph, dataDir & "TestFile.RenderParagraph Out.png", options)
			'ExEnd
		End Sub

		Public Shared Sub FindShapeSizes(ByVal shape As Shape)
			'ExStart
			'ExFor:ShapeRenderer.SizeInPoints
			'ExId:ShapeRendererSizeInPoints
			'ExSummary:Demonstrates how to find the size of a shape in the document and the size of the shape when rendered.
			Dim shapeSizeInDocument As SizeF = shape.GetShapeRenderer().SizeInPoints
			Dim width As Single = shapeSizeInDocument.Width ' The width of the shape.
			Dim height As Single = shapeSizeInDocument.Height ' The height of the shape.
			'ExEnd

			'ExStart
			'ExFor:ShapeRenderer.GetSizeInPixels
			'ExId:ShapeRendererGetSizeInPixels
			'ExSummary:Shows how to create a new Bitmap and Graphics object with the width and height of the shape to be rendered.
			' We will render the shape at normal size and 96dpi. Calculate the size in pixels that the shape will be rendered at.
			Dim shapeRenderedSize As Size = shape.GetShapeRenderer().GetSizeInPixels(1.0f, 96.0f)

			Using image As New Bitmap(shapeRenderedSize.Width, shapeRenderedSize.Height)
				Using g As Graphics = Graphics.FromImage(image)
					' Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.
				End Using
			End Using
			'ExEnd
		End Sub

		'ExStart
		'ExId:RenderNode
		'ExSummary:Shows how to render a node independent of the document by building on the functionality provided by ShapeRenderer class.
		''' <summary>
		''' Renders any node in a document to the path specified using the image save options.
		''' </summary>
		''' <param name="node">The node to render.</param>
		''' <param name="path">The path to save the rendered image to.</param>
		''' <param name="imageOptions">The image options to use during rendering. This can be null.</param>
		Public Shared Sub RenderNode(ByVal node As Node, ByVal filePath As String, ByVal imageOptions As ImageSaveOptions)
			' Run some argument checks.
			If node Is Nothing Then
				Throw New ArgumentException("Node cannot be null")
			End If

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
			Dim shape As New Shape(doc, ShapeType.TextBox)
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
			Do While Not(TypeOf currentNode.ParentNode Is InlineStory OrElse TypeOf currentNode.ParentNode Is Story OrElse TypeOf currentNode.ParentNode Is ShapeBase)
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
			Dim stream As New MemoryStream()
			shape.GetShapeRenderer().Save(stream, imageOptions)
			shape.Remove()

			' Load the image into a new bitmap.
			Using renderedImage As New Bitmap(stream)
				' Extract the actual content of the image by cropping transparent space around
				' the rendered shape.
				Dim cropRectangle As Rectangle = FindBoundingBoxAroundNode(renderedImage)

				Dim croppedImage As New Bitmap(cropRectangle.Width, cropRectangle.Height)
				croppedImage.SetResolution(imageOptions.Resolution, imageOptions.Resolution)

				' Create the final image with the proper background color.
				Using g As Graphics = Graphics.FromImage(croppedImage)
					g.Clear(savePaperColor)
					g.DrawImage(renderedImage, New Rectangle(0, 0, croppedImage.Width, croppedImage.Height), cropRectangle.X, cropRectangle.Y, cropRectangle.Width, cropRectangle.Height, GraphicsUnit.Pixel)
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
						min.X = Math.Min(x, min.X)
						min.Y = Math.Min(y, min.Y)
						max.X = Math.Max(x, max.X)
						max.Y = Math.Max(y, max.Y)
					End If
				Next y
			Next x

			' Add one pixel to the width and height to avoid clipping.
			Return New Rectangle(min.X, min.Y, (max.X - min.X) + 1, (max.Y - min.Y) + 1)
		End Function
		'ExEnd
	End Class
End Namespace