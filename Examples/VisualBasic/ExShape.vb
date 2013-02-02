'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports NUnit.Framework
Imports System

Namespace Examples
	''' <summary>
	''' Examples using shapes in documents.
	''' </summary>
	<TestFixture> _
	Public Class ExShape
		Inherits ExBase
		<Test> _
		Public Sub DeleteAllShapes()
			Dim doc As New Document(MyDir & "Shape.DeleteAllShapes.doc")

			'ExStart
			'ExFor:Shape
			'ExSummary:Shows how to delete all shapes from a document.
			' Here we get all shapes from the document node, but you can do this for any smaller
			' node too, for example delete shapes from a single section or a paragraph.
			Dim shapes As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)
			shapes.Clear()

			' There could also be group shapes, they have different node type, remove them all too.
			Dim groupShapes As NodeCollection = doc.GetChildNodes(NodeType.GroupShape, True)
			groupShapes.Clear()
			'ExEnd

			Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, True).Count)
			Assert.AreEqual(0, doc.GetChildNodes(NodeType.GroupShape, True).Count)
			doc.Save(MyDir & "Shape.DeleteAllShapes Out.doc")
		End Sub

		<Test> _
		Public Sub CheckShapeInline()
			'ExStart
			'ExFor:ShapeBase.IsInline
			'ExSummary:Shows how to test if a shape in the document is inline or floating.
			Dim doc As New Document(MyDir & "Shape.DeleteAllShapes.doc")

			For Each shape As Shape In doc.GetChildNodes(NodeType.Shape, True)
				If shape.IsInline Then
					Console.WriteLine("Shape is inline.")
				Else
					Console.WriteLine("Shape is floating.")
				End If
			Next shape

			'ExEnd

			' Verify that the first shape in the document is not inline.
			Assert.False((CType(doc.GetChild(NodeType.Shape, 0, True), Shape)).IsInline)
		End Sub

		<Test> _
		Public Sub LineFlipOrientation()
			'ExStart
			'ExFor:ShapeBase.Bounds
			'ExFor:ShapeBase.FlipOrientation
			'ExFor:FlipOrientation
			'ExSummary:Creates two line shapes. One line goes from top left to bottom right. Another line goes from bottom left to top right.
			Dim doc As New Document()

			' The lines will cross the whole page.
			Dim pageWidth As Single = CSng(doc.FirstSection.PageSetup.PageWidth)
			Dim pageHeight As Single= CSng(doc.FirstSection.PageSetup.PageHeight)

			' This line goes from top left to bottom right by default. 
			Dim lineA As New Shape(doc, ShapeType.Line)
			lineA.Bounds = New RectangleF(0, 0, pageWidth, pageHeight)
			lineA.RelativeHorizontalPosition = RelativeHorizontalPosition.Page
			lineA.RelativeVerticalPosition = RelativeVerticalPosition.Page
			doc.FirstSection.Body.FirstParagraph.AppendChild(lineA)

			' This line goes from bottom left to top right because we flipped it. 
			Dim lineB As New Shape(doc, ShapeType.Line)
			lineB.Bounds = New RectangleF(0, 0, pageWidth, pageHeight)
			lineB.FlipOrientation = FlipOrientation.Horizontal
			lineB.RelativeHorizontalPosition = RelativeHorizontalPosition.Page
			lineB.RelativeVerticalPosition = RelativeVerticalPosition.Page
			doc.FirstSection.Body.FirstParagraph.AppendChild(lineB)

			doc.Save(MyDir & "Shape.LineFlipOrientation Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub Fill()
			'ExStart
			'ExFor:Shape.Fill
			'ExFor:Shape.FillColor
			'ExFor:Fill
			'ExFor:Fill.Opacity
			'ExSummary:Demonstrates how to create shapes with fill.
			Dim builder As New DocumentBuilder()

			builder.Writeln()
			builder.Writeln()
			builder.Writeln()
			builder.Write("Some text under the shape.")

			' Create a red balloon, semitransparent.
			' The shape is floating and its coordinates are (0,0) by default, relative to the current paragraph.
			Dim shape As New Shape(builder.Document, ShapeType.Balloon)
			shape.FillColor = Color.Red
			shape.Fill.Opacity = 0.3
			shape.Width = 100
			shape.Height = 100
			shape.Top = -100
			builder.InsertNode(shape)

			builder.Document.Save(MyDir & "Shape.Fill Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ReplaceTextboxesWithImages()
			'ExStart
			'ExFor:WrapSide
			'ExFor:ShapeBase.WrapSide
			'ExFor:NodeCollection
			'ExFor:CompositeNode.InsertAfter(Node, Node)
			'ExFor:NodeCollection.ToArray
			'ExSummary:Shows how to replace all textboxes with images.
			Dim doc As New Document(MyDir & "Shape.ReplaceTextboxesWithImages.doc")

			' This gets a live collection of all shape nodes in the document.
			Dim shapeCollection As NodeCollection = doc.GetChildNodes(NodeType.Shape, True)

			' Since we will be adding/removing nodes, it is better to copy all collection
			' into a fixed size array, otherwise iterator will be invalidated.
			Dim shapes() As Node = shapeCollection.ToArray()

			For Each shape As Shape In shapes
				' Filter out all shapes that we don't need.
				If shape.ShapeType.Equals(ShapeType.TextBox) Then
					' Create a new shape that will replace the existing shape.
					Dim image As New Shape(doc, ShapeType.Image)

					' Load the image into the new shape.
					image.ImageData.SetImage(MyDir & "Hammer.wmf")

					' Make new shape's position to match the old shape.
					image.Left = shape.Left
					image.Top = shape.Top
					image.Width = shape.Width
					image.Height = shape.Height
					image.RelativeHorizontalPosition = shape.RelativeHorizontalPosition
					image.RelativeVerticalPosition = shape.RelativeVerticalPosition
					image.HorizontalAlignment = shape.HorizontalAlignment
					image.VerticalAlignment = shape.VerticalAlignment
					image.WrapType = shape.WrapType
					image.WrapSide = shape.WrapSide

					' Insert new shape after the old shape and remove the old shape.
					shape.ParentNode.InsertAfter(image, shape)
					shape.Remove()
				End If
			Next shape

			doc.Save(MyDir & "Shape.ReplaceTextboxesWithImages Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub CreateTextBox()
			'ExStart
			'ExFor:Shape.#ctor(DocumentBase, ShapeType)
			'ExFor:ShapeBase.ZOrder
			'ExFor:Story.FirstParagraph
			'ExFor:Shape.FirstParagraph
			'ExFor:ShapeBase.WrapType
			'ExSummary:Creates a textbox with some text and different formatting options in a new document.
			' Create a blank document.
			Dim doc As New Document()

			' Create a new shape of type TextBox
			Dim textBox As New Shape(doc, ShapeType.TextBox)

			' Set some settings of the textbox itself.
			' Set the wrap of the textbox to inline
			textBox.WrapType = WrapType.None
			' Set the horizontal and vertical alignment of the text inside the shape.
			textBox.HorizontalAlignment = HorizontalAlignment.Center
			textBox.VerticalAlignment = VerticalAlignment.Top

			' Set the textbox height and width.
			textBox.Height = 50
			textBox.Width = 200

			' Set the textbox in front of other shapes with a lower ZOrder
			textBox.ZOrder = 2

			' Let's create a new paragraph for the textbox manually and align it in the center. Make sure we add the new nodes to the textbox as well.
			textBox.AppendChild(New Paragraph(doc))
			Dim para As Paragraph = textBox.FirstParagraph
			para.ParagraphFormat.Alignment = ParagraphAlignment.Center

			' Add some text to the paragraph.
			Dim run As New Run(doc)
			run.Text = "Content in textbox"
			para.AppendChild(run)

			' Append the textbox to the first paragraph in the body.
			doc.FirstSection.Body.FirstParagraph.AppendChild(textBox)

			' Save the output
			doc.Save(MyDir & "Shape.CreateTextBox Out.doc")
			'ExEnd
		End Sub
	End Class
End Namespace
