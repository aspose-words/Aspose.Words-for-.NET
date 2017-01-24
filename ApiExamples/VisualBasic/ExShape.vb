' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Ole
Imports Aspose.Words.Math
Imports Aspose.Words.Rendering
Imports Aspose.Words.Saving

Imports NUnit.Framework

Namespace ApiExamples
	''' <summary>
	''' Examples using shapes in documents.
	''' </summary>
	<TestFixture> _
	Public Class ExShape
		Inherits ApiExampleBase
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
			doc.Save(MyDir & "\Artifacts\Shape.DeleteAllShapes.doc")
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

			doc.Save(MyDir & "\Artifacts\Shape.LineFlipOrientation.doc")
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

			builder.Document.Save(MyDir & "\Artifacts\Shape.Fill.doc")
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

			doc.Save(MyDir & "\Artifacts\Shape.ReplaceTextboxesWithImages.doc")
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
			doc.Save(MyDir & "\Artifacts\Shape.CreateTextBox.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetActiveXControlProperties()
			'ExStart
			'ExFor:OleControl
			'ExFor:Forms2OleControlCollection.Caption
			'ExFor:Forms2OleControlCollection.Value
			'ExFor:Forms2OleControlCollection.Enabled
			'ExFor:Forms2OleControlCollection.Type
			'ExFor:Forms2OleControlCollection.ChildNodes
			'ExSummary: Shows how to get ActiveX control and properties from the document
			Dim doc As New Document(MyDir & "Shape.ActiveXObject.docx")

			'Get ActiveX control from the document 
			Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
			Dim oleControl As OleControl = shape.OleFormat.OleControl

			'Get ActiveX control properties
			If oleControl.IsForms2OleControl Then
				Dim checkBox As Forms2OleControl = CType(oleControl, Forms2OleControl)
				Assert.AreEqual("Первый", checkBox.Caption)
				Assert.AreEqual("0", checkBox.Value)
				Assert.AreEqual(True, checkBox.Enabled)
				Assert.AreEqual(Forms2OleControlType.CheckBox, checkBox.Type)
				Assert.AreEqual(Nothing, checkBox.ChildNodes)
			End If
			'ExEnd
		End Sub

		<Test> _
		Public Sub SuggestedFileName()
			'ExStart
			'ExFor:OleFormat.SuggestedFileName
			'ExSummary:Shows how to get suggested file name from the object
			Dim doc As New Document(MyDir & "Shape.SuggestedFileName.rtf")

			'Gets the file name suggested for the current embedded object if you want to save it into a file
			Dim oleShape As Shape = CType(doc.FirstSection.Body.GetChild(NodeType.Shape, 0, True), Shape)
			Dim suggestedFileName As String = oleShape.OleFormat.SuggestedFileName
			'ExEnd

			Assert.AreEqual("CSV.csv", suggestedFileName)
		End Sub

		<Test> _
		Public Sub ObjectDidNotHaveSuggestedFileName()
			Dim doc As New Document(MyDir & "Shape.ActiveXObject.docx")

			Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
			Assert.IsEmpty(shape.OleFormat.SuggestedFileName)
		End Sub

		<Test> _
		Public Sub GetOpaqueBoundsInPixels()
			Dim doc As New Document(MyDir & "Shape.TextBox.doc")

			Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)

			Dim imageOptions As New ImageSaveOptions(SaveFormat.Jpeg)

			Dim stream As New MemoryStream()
			Dim renderer As ShapeRenderer = shape.GetShapeRenderer()
			renderer.Save(stream, imageOptions)
			shape.Remove()

			'Check that the opaque bounds and bounds have default values
			Assert.AreEqual(250, renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.Resolution).Width)
			Assert.AreEqual(52, renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.Resolution).Height)

			Assert.AreEqual(250, renderer.GetBoundsInPixels(imageOptions.Scale, imageOptions.Resolution).Width)
			Assert.AreEqual(52, renderer.GetBoundsInPixels(imageOptions.Scale, imageOptions.Resolution).Height)
		End Sub

		'For assert result of the test you need to open "Document.OfficeMath Out.svg" and check that OfficeMath node is there
		<Test> _
		Public Sub SaveShapeObjectAsImage()
			'ExStart
			'ExFor:Shows how to convert specific object into image
			Dim doc As New Document(MyDir & "Document.OfficeMath.docx")

			'Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
			Dim math As OfficeMath = CType(doc.GetChild(NodeType.OfficeMath, 0, True), OfficeMath)
			math.GetMathRenderer().Save(MyDir & "Document.OfficeMath Out.svg", New ImageSaveOptions(SaveFormat.Svg))
			'ExEnd
		End Sub

		<Test, TestCase(True), TestCase(False)> _
		Public Sub AspectRatioLocked(ByVal isLocked As Boolean)
			'ExStart
			'ExFor:Shows how to set "AspectRatioLocked" for the shape object
			Dim doc As New Document(MyDir & "Shape.ActiveXObject.docx")

			'Get shape object from the document and set AspectRatioLocked(it is possible to get/set AspectRatioLocked for child shapes (mimic MS Word behavior), but AspectRatioLocked has effect only for top level shapes!)
			Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
			shape.AspectRatioLocked = isLocked
			'ExEnd

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
			Assert.AreEqual(isLocked, shape.AspectRatioLocked)
		End Sub
	End Class
End Namespace
