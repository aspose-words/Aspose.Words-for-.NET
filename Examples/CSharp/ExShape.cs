//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;
using System;

namespace Examples
{
    /// <summary>
    /// Examples using shapes in documents.
    /// </summary>
    [TestFixture]
    public class ExShape : ExBase
    {
        [Test]
        public void DeleteAllShapes()
        {
            Document doc = new Document(MyDir + "Shape.DeleteAllShapes.doc");
            
            //ExStart
            //ExFor:Shape
            //ExSummary:Shows how to delete all shapes from a document.
            // Here we get all shapes from the document node, but you can do this for any smaller
            // node too, for example delete shapes from a single section or a paragraph.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true, false);
            shapes.Clear();

            // There could also be group shapes, they have different node type, remove them all too.
            NodeCollection groupShapes = doc.GetChildNodes(NodeType.GroupShape, true, false);
            groupShapes.Clear();
            //ExEnd

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true, false).Count);
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.GroupShape, true, false).Count);
            doc.Save(MyDir + "Shape.DeleteAllShapes Out.doc");
        }

        [Test]
        public void CheckShapeInline()
        {
            //ExStart
            //ExFor:ShapeBase.IsInline
            //ExSummary:Shows how to test if a shape in the document is inline or floating.
            Document doc = new Document(MyDir + "Shape.DeleteAllShapes.doc");

            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                if(shape.IsInline)
                    Console.WriteLine("Shape is inline.");
                else
                    Console.WriteLine("Shape is floating.");
            }

            //ExEnd
            
            // Verify that the first shape in the document is not inline.
            Assert.False(((Shape)doc.GetChild(NodeType.Shape, 0, true)).IsInline);
        }

        [Test]
        public void LineFlipOrientation()
        {
            //ExStart
            //ExFor:ShapeBase.Bounds
            //ExFor:ShapeBase.FlipOrientation
            //ExFor:FlipOrientation
            //ExSummary:Creates two line shapes. One line goes from top left to bottom right. Another line goes from bottom left to top right.
            Document doc = new Document();

            // The lines will cross the whole page.
            float pageWidth = (float)doc.FirstSection.PageSetup.PageWidth;
            float pageHeight= (float)doc.FirstSection.PageSetup.PageHeight;

            // This line goes from top left to bottom right by default. 
            Shape lineA = new Shape(doc, ShapeType.Line);
            lineA.Bounds = new RectangleF(0, 0, pageWidth, pageHeight);
            lineA.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            lineA.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            doc.FirstSection.Body.FirstParagraph.AppendChild(lineA);

            // This line goes from bottom left to top right because we flipped it. 
            Shape lineB = new Shape(doc, ShapeType.Line);
            lineB.Bounds = new RectangleF(0, 0, pageWidth, pageHeight);
            lineB.FlipOrientation = FlipOrientation.Horizontal;
            lineB.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            lineB.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            doc.FirstSection.Body.FirstParagraph.AppendChild(lineB);

            doc.Save(MyDir + "Shape.LineFlipOrientation Out.doc");
            //ExEnd
        }

        [Test]
        public void Fill()
        {
            //ExStart
            //ExFor:Shape.Fill
            //ExFor:Shape.FillColor
            //ExFor:Fill
            //ExFor:Fill.Opacity
            //ExSummary:Demonstrates how to create shapes with fill.
            DocumentBuilder builder = new DocumentBuilder();

            builder.Writeln();
            builder.Writeln();
            builder.Writeln();
            builder.Write("Some text under the shape.");

            // Create a red balloon, semitransparent.
            // The shape is floating and its coordinates are (0,0) by default, relative to the current paragraph.
            Shape shape = new Shape(builder.Document, ShapeType.Balloon);
            shape.FillColor = Color.Red;
            shape.Fill.Opacity = 0.3;
            shape.Width = 100;
            shape.Height = 100;
            shape.Top = -100;
            builder.InsertNode(shape);

            builder.Document.Save(MyDir + "Shape.Fill Out.doc");
            //ExEnd
        }

        [Test]
        public void ReplaceTextboxesWithImages()
        {
            //ExStart
            //ExFor:WrapSide
            //ExFor:ShapeBase.WrapSide
            //ExFor:NodeCollection
            //ExFor:CompositeNode.InsertAfter(Node, Node)
            //ExFor:NodeCollection.ToArray
            //ExSummary:Shows how to replace all textboxes with images.
            Document doc = new Document(MyDir + "Shape.ReplaceTextboxesWithImages.doc");

            // This gets a live collection of all shape nodes in the document.
            NodeCollection shapeCollection = doc.GetChildNodes(NodeType.Shape, true);

            // Since we will be adding/removing nodes, it is better to copy all collection
            // into a fixed size array, otherwise iterator will be invalidated.
            Node[] shapes = shapeCollection.ToArray();

            foreach (Shape shape in shapes)
            {
                // Filter out all shapes that we don't need.
                if (shape.ShapeType.Equals(ShapeType.TextBox))
                {
                    // Create a new shape that will replace the existing shape.
                    Shape image = new Shape(doc, ShapeType.Image);

                    // Load the image into the new shape.
                    image.ImageData.SetImage(MyDir + "Hammer.wmf");

                    // Make new shape's position to match the old shape.
                    image.Left = shape.Left;
                    image.Top = shape.Top;
                    image.Width = shape.Width;
                    image.Height = shape.Height;
                    image.RelativeHorizontalPosition = shape.RelativeHorizontalPosition;
                    image.RelativeVerticalPosition = shape.RelativeVerticalPosition;
                    image.HorizontalAlignment = shape.HorizontalAlignment;
                    image.VerticalAlignment = shape.VerticalAlignment;
                    image.WrapType = shape.WrapType;
                    image.WrapSide = shape.WrapSide;

                    // Insert new shape after the old shape and remove the old shape.
                    shape.ParentNode.InsertAfter(image, shape);
                    shape.Remove();
                }
            }

            doc.Save(MyDir + "Shape.ReplaceTextboxesWithImages Out.doc");
            //ExEnd
        }

        [Test]
        public void CreateTextBox()
        {
            //ExStart
            //ExFor:Shape.#ctor(DocumentBase, ShapeType)
            //ExFor:ShapeBase.ZOrder
            //ExFor:Story.FirstParagraph
            //ExFor:Shape.FirstParagraph
            //ExFor:ShapeBase.WrapType
            //ExSummary:Creates a textbox with some text and different formatting options in a new document.
            // Create a blank document.
            Document doc = new Document();

            // Create a new shape of type TextBox
            Shape textBox = new Shape(doc, ShapeType.TextBox);

            // Set some settings of the textbox itself.
            // Set the wrap of the textbox to inline
            textBox.WrapType = WrapType.None;
            // Set the horizontal and vertical alignment of the text inside the shape.
            textBox.HorizontalAlignment = HorizontalAlignment.Center;
            textBox.VerticalAlignment = VerticalAlignment.Top;

            // Set the textbox height and width.
            textBox.Height = 50;
            textBox.Width = 200;

            // Set the textbox in front of other shapes with a lower ZOrder
            textBox.ZOrder = 2;

            // Let's create a new paragraph for the textbox manually and align it in the center. Make sure we add the new nodes to the textbox as well.
            textBox.AppendChild(new Paragraph(doc));
            Paragraph para = textBox.FirstParagraph;
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Add some text to the paragraph.
            Run run = new Run(doc);
            run.Text = "Content in textbox";
            para.AppendChild(run);

            // Append the textbox to the first paragraph in the body.
            doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

            // Save the output
            doc.Save(MyDir + "Shape.CreateTextBox Out.doc");
            //ExEnd
        }
    }
}
