' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.Drawing
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Drawing

Imports NUnit.Framework

Namespace ApiExamples
	''' <summary>
	''' Mostly scenarios that deal with image shapes.
	''' </summary>
	<TestFixture> _
	Public Class ExImage
		Inherits ApiExampleBase
		<Test> _
		Public Sub CreateFromUrl()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(string)
			'ExFor:DocumentBuilder.Writeln
			'ExSummary:Shows how to inserts an image from a URL. The image is inserted inline and at 100% scale.
			' This creates a builder and also an empty document inside the builder.
			Dim builder As New DocumentBuilder()

			builder.Write("Image from local file: ")
			builder.InsertImage(MyDir & "Aspose.Words.gif")
			builder.Writeln()

			builder.Write("Image from an internet url, automatically downloaded for you: ")
			builder.InsertImage("http://www.aspose.com/Images/aspose-logo.jpg")
			builder.Writeln()

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromUrl.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CreateFromStream()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Stream)
			'ExSummary:Shows how to insert an image from a stream. The image is inserted inline and at 100% scale.
			' This creates a builder and also an empty document inside the builder.
			Dim builder As New DocumentBuilder()

			Dim stream As Stream = File.OpenRead(MyDir & "Aspose.Words.gif")
			Try
				builder.Write("Image from stream: ")
				builder.InsertImage(stream)
			Finally
				stream.Close()
			End Try

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromStream.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CreateFromImage()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(Image)
			'ExSummary:Shows how to insert a .NET Image object into a document. The image is inserted inline and at 100% scale.
			' This creates a builder and also an empty document inside the builder.
			Dim builder As New DocumentBuilder()

			' Insert a raster image.
			Dim rasterImage As Image = Image.FromFile(MyDir & "Aspose.Words.gif")
			Try
				builder.Write("Raster image: ")
				builder.InsertImage(rasterImage)
				builder.Writeln()
			Finally
				rasterImage.Dispose()
			End Try

			' Aspose.Words allows to insert a metafile too.
			Dim metafile As Image = Image.FromFile(MyDir & "Hammer.wmf")
			Try
				builder.Write("Metafile: ")
				builder.InsertImage(metafile)
				builder.Writeln()
			Finally
				metafile.Dispose()
			End Try

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFromImage.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CreateFloatingPageCenter()
			'ExStart
			'ExFor:DocumentBuilder.InsertImage(string)
			'ExFor:Shape
			'ExFor:ShapeBase
			'ExFor:ShapeBase.WrapType
			'ExFor:ShapeBase.BehindText
			'ExFor:ShapeBase.RelativeHorizontalPosition
			'ExFor:ShapeBase.RelativeVerticalPosition
			'ExFor:ShapeBase.HorizontalAlignment
			'ExFor:ShapeBase.VerticalAlignment
			'ExFor:WrapType
			'ExFor:RelativeHorizontalPosition
			'ExFor:RelativeVerticalPosition
			'ExFor:HorizontalAlignment
			'ExFor:VerticalAlignment
			'ExSummary:Shows how to insert a floating image in the middle of a page.
			' This creates a builder and also an empty document inside the builder.
			Dim builder As New DocumentBuilder()

			' By default, the image is inline.
			Dim shape As Shape = builder.InsertImage(MyDir & "Aspose.Words.gif")

			' Make the image float, put it behind text and center on the page.
			shape.WrapType = WrapType.None
			shape.BehindText = True
			shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page
			shape.HorizontalAlignment = HorizontalAlignment.Center
			shape.RelativeVerticalPosition = RelativeVerticalPosition.Page
			shape.VerticalAlignment = VerticalAlignment.Center

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFloatingPageCenter.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CreateFloatingPositionSize()
			'ExStart
			'ExFor:ShapeBase.Left
			'ExFor:ShapeBase.Top
			'ExFor:ShapeBase.Width
			'ExFor:ShapeBase.Height
			'ExFor:DocumentBuilder.CurrentSection
			'ExFor:PageSetup.PageWidth
			'ExSummary:Shows how to insert a floating image and specify its position and size.
			' This creates a builder and also an empty document inside the builder.
			Dim builder As New DocumentBuilder()

			' By default, the image is inline.
			Dim shape As Shape = builder.InsertImage(MyDir & "Hammer.wmf")

			' Make the image float, put it behind text and center on the page.
			shape.WrapType = WrapType.None

			' Make position relative to the page.
			shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page
			shape.RelativeVerticalPosition = RelativeVerticalPosition.Page

			' Make the shape occupy a band 50 points high at the very top of the page.
			shape.Left = 0
			shape.Top = 0
			shape.Width = builder.CurrentSection.PageSetup.PageWidth
			shape.Height = 50

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateFloatingPositionSize.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertImageWithHyperlink()
			'ExStart
			'ExFor:ShapeBase.HRef
			'ExFor:ShapeBase.ScreenTip
			'ExSummary:Shows how to insert an image with a hyperlink.
			' This creates a builder and also an empty document inside the builder.
			Dim builder As New DocumentBuilder()

			Dim shape As Shape = builder.InsertImage(MyDir & "Hammer.wmf")
			shape.HRef = "http://www.aspose.com/Community/Forums/75/ShowForum.aspx"
			shape.ScreenTip = "Aspose.Words Support Forums"

			builder.Document.Save(MyDir & "\Artifacts\Image.InsertImageWithHyperlink.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CreateImageDirectly()
			'ExStart
			'ExFor:Shape.#ctor(DocumentBase,ShapeType)
			'ExFor:ShapeType
			'ExSummary:Shows how to create and add an image to a document without using document builder.
			Dim doc As New Document()

			Dim shape As New Shape(doc, ShapeType.Image)
			shape.ImageData.SetImage(MyDir & "Hammer.wmf")
			shape.Width = 100
			shape.Height = 100

			doc.FirstSection.Body.FirstParagraph.AppendChild(shape)

			doc.Save(MyDir & "\Artifacts\Image.CreateImageDirectly.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CreateLinkedImage()
			'ExStart
			'ExFor:Shape.ImageData
			'ExFor:ImageData
			'ExFor:ImageData.SourceFullName
			'ExFor:ImageData.SetImage(string)
			'ExFor:DocumentBuilder.InsertNode
			'ExSummary:Shows how to insert a linked image into a document. 
			Dim builder As New DocumentBuilder()

			Dim imageFileName As String = MyDir & "Hammer.wmf"

			builder.Write("Image linked, not stored in the document: ")

			Dim linkedOnly As New Shape(builder.Document, ShapeType.Image)
			linkedOnly.WrapType = WrapType.Inline
			linkedOnly.ImageData.SourceFullName = imageFileName

			builder.InsertNode(linkedOnly)
			builder.Writeln()


			builder.Write("Image linked and stored in the document: ")

			Dim linkedAndStored As New Shape(builder.Document, ShapeType.Image)
			linkedAndStored.WrapType = WrapType.Inline
			linkedAndStored.ImageData.SourceFullName = imageFileName
			linkedAndStored.ImageData.SetImage(imageFileName)

			builder.InsertNode(linkedAndStored)
			builder.Writeln()


			builder.Write("Image stored in the document, but not linked: ")

			Dim stored As New Shape(builder.Document, ShapeType.Image)
			stored.WrapType = WrapType.Inline
			stored.ImageData.SetImage(imageFileName)

			builder.InsertNode(stored)
			builder.Writeln()

			builder.Document.Save(MyDir & "\Artifacts\Image.CreateLinkedImage.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub DeleteAllImages()
			Dim doc As New Document(MyDir & "Image.SampleImages.doc")
			Assert.AreEqual(6, doc.GetChildNodes(NodeType.Shape, True).Count)

			'ExStart
			'ExFor:Shape.HasImage
			'ExFor:Node.Remove
			'ExSummary:Shows how to delete all images from a document.
			' Here we get all shapes from the document node, but you can do this for any smaller
			' node too, for example delete shapes from a single section or a paragraph.
			Dim shapes As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)

			' We cannot delete shape nodes while we enumerate through the collection.
			' One solution is to add nodes that we want to delete to a temporary array and delete afterwards.
			Dim shapesToDelete As New ArrayList()
			For Each shape As Shape In shapes
				' Several shape types can have an image including image shapes and OLE objects.
				If shape.HasImage Then
					shapesToDelete.Add(shape)
				End If
			Next shape

			' Now we can delete shapes.
			For Each shape As Shape In shapesToDelete
				shape.Remove()
			Next shape
			'ExEnd

			Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, True).Count)
			doc.Save(MyDir & "\Artifacts\Image.DeleteAllImages.doc")
		End Sub

		<Test> _
		Public Sub DeleteAllImagesPreOrder()
			Dim doc As New Document(MyDir & "Image.SampleImages.doc")
			Assert.AreEqual(6, doc.GetChildNodes(NodeType.Shape, True).Count)

			'ExStart
			'ExFor:Node.NextPreOrder
			'ExSummary:Shows how to delete all images from a document using pre-order tree traversal.
			Dim curNode As Node = doc
			Do While curNode IsNot Nothing
				Dim nextNode As Node = curNode.NextPreOrder(doc)

				If curNode.NodeType.Equals(NodeType.Shape) Then
					Dim shape As Shape = CType(curNode, Shape)

					' Several shape types can have an image including image shapes and OLE objects.
					If shape.HasImage Then
						shape.Remove()
					End If
				End If

				curNode = nextNode
			Loop
			'ExEnd

			Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, True).Count)
			doc.Save(MyDir & "\Artifacts\Image.DeleteAllImagesPreOrder.doc")
		End Sub

		'ExStart
		'ExFor:Shape
		'ExFor:Shape.ImageData
		'ExFor:Shape.HasImage
		'ExFor:ImageData
		'ExFor:FileFormatUtil.ImageTypeToExtension(Aspose.Words.Drawing.ImageType)
		'ExFor:ImageData.ImageType
		'ExFor:ImageData.Save(string)
		'ExFor:CompositeNode.GetChildNodes(NodeType, bool)
		'ExId:ExtractImagesToFiles
		'ExSummary:Shows how to extract images from a document and save them as files.
		<Test> _
		Public Sub ExtractImagesToFiles()
			Dim doc As New Document(MyDir & "Image.SampleImages.doc")

			Dim shapes As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)
			Dim imageIndex As Integer = 0
			For Each shape As Shape In shapes
				If shape.HasImage Then
					Dim imageFileName As String = String.Format("\Artifacts\Image.ExportImages.{0} Out{1}", imageIndex, FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType))
					shape.ImageData.Save(MyDir & imageFileName)
					imageIndex += 1
				End If
			Next shape
		End Sub
		'ExEnd

		<Test> _
		Public Sub ScaleImage()
			'ExStart
			'ExFor:ImageData.ImageSize
			'ExFor:ImageSize
			'ExFor:ImageSize.WidthPoints
			'ExFor:ImageSize.HeightPoints
			'ExFor:ShapeBase.Width
			'ExFor:ShapeBase.Height
			'ExSummary:Shows how to resize an image shape.
			Dim builder As New DocumentBuilder()

			' By default, the image is inserted at 100% scale.
			Dim shape As Shape = builder.InsertImage(MyDir & "Aspose.Words.gif")

			' It is easy to change the shape size. In this case, make it 50% relative to the current shape size.
			shape.Width = shape.Width * 0.5
			shape.Height = shape.Height * 0.5

			' However, we can also go back to the original image size and scale from there, say 110%.
			Dim imageSize As ImageSize = shape.ImageData.ImageSize
			shape.Width = imageSize.WidthPoints * 1.1
			shape.Height = imageSize.HeightPoints * 1.1

			builder.Document.Save(MyDir & "\Artifacts\Image.ScaleImage.doc")
			'ExEnd
		End Sub
	End Class
End Namespace
