// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.IO;

using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;
using Aspose.Words.Math;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Examples using shapes in documents.
    /// </summary>
    [TestFixture]
    public class ExShape : ApiExampleBase
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
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            shapes.Clear();

            // There could also be group shapes, they have different node type, remove them all too.
            NodeCollection groupShapes = doc.GetChildNodes(NodeType.GroupShape, true);
            groupShapes.Clear();
            //ExEnd

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true).Count);
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.GroupShape, true).Count);
            doc.Save(MyDir + @"\Artifacts\Shape.DeleteAllShapes.doc");
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

            doc.Save(MyDir + @"\Artifacts\Shape.LineFlipOrientation.doc");
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

            builder.Document.Save(MyDir + @"\Artifacts\Shape.Fill.doc");
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

            doc.Save(MyDir + @"\Artifacts\Shape.ReplaceTextboxesWithImages.doc");
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
            doc.Save(MyDir + @"\Artifacts\Shape.CreateTextBox.doc");
            //ExEnd
        }

        [Test]
        public void GetActiveXControlProperties()
        {
            //ExStart
            //ExFor:OleControl
            //ExFor:Forms2OleControlCollection.Caption
            //ExFor:Forms2OleControlCollection.Value
            //ExFor:Forms2OleControlCollection.Enabled
            //ExFor:Forms2OleControlCollection.Type
            //ExFor:Forms2OleControlCollection.ChildNodes
            //ExSummary: Shows how to get ActiveX control and properties from the document
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");
            
            //Get ActiveX control from the document 
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            OleControl oleControl = shape.OleFormat.OleControl;

            //Get ActiveX control properties
            if (oleControl.IsForms2OleControl)
            {
                Forms2OleControl checkBox = (Forms2OleControl)oleControl;
                Assert.AreEqual("Первый", checkBox.Caption);
                Assert.AreEqual("0", checkBox.Value);
                Assert.AreEqual(true, checkBox.Enabled);
                Assert.AreEqual(Forms2OleControlType.CheckBox, checkBox.Type);
                Assert.AreEqual(null, checkBox.ChildNodes);
            }
            //ExEnd
        }

        [Test]
        public void SuggestedFileName()
        {
            //ExStart
            //ExFor:OleFormat.SuggestedFileName
            //ExSummary:Shows how to get suggested file name from the object
            Document doc = new Document(MyDir + "Shape.SuggestedFileName.rtf");
            
            //Gets the file name suggested for the current embedded object if you want to save it into a file
            Shape oleShape = (Shape)doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true);
            string suggestedFileName = oleShape.OleFormat.SuggestedFileName;
            //ExEnd

            Assert.AreEqual("CSV.csv", suggestedFileName);
        }

        [Test]
        public void ObjectDidNotHaveSuggestedFileName()
        {
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.IsEmpty(shape.OleFormat.SuggestedFileName);
        }

        [Test]
        public void GetOpaqueBoundsInPixels()
        {
            Document doc = new Document(MyDir + "Shape.TextBox.doc");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            MemoryStream stream = new MemoryStream();
            ShapeRenderer renderer = shape.GetShapeRenderer();
            renderer.Save(stream, imageOptions);
            shape.Remove();

            //Check that the opaque bounds and bounds have default values
            Assert.AreEqual(250, renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.Resolution).Width);
            Assert.AreEqual(52, renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.Resolution).Height);

            Assert.AreEqual(250, renderer.GetBoundsInPixels(imageOptions.Scale, imageOptions.Resolution).Width);
            Assert.AreEqual(52, renderer.GetBoundsInPixels(imageOptions.Scale, imageOptions.Resolution).Height);

            Assert.AreEqual((float)187.849991, renderer.OpaqueBoundsInPoints.Width);
            Assert.AreEqual((float)39.25, renderer.OpaqueBoundsInPoints.Height);
        }

        //For assert result of the test you need to open "Document.OfficeMath Out.svg" and check that OfficeMath node is there
        [Test]
        public void SaveShapeObjectAsImage()
        {
            //ExStart
            //ExFor:OfficeMath.GetMathRenderer.Save
            //ExSummary:Shows how to convert specific object into image
            Document doc = new Document(MyDir + "Document.OfficeMath.docx");

            //Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
            OfficeMath math = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            math.GetMathRenderer().Save(MyDir + "Document.OfficeMath Out.svg", new ImageSaveOptions(SaveFormat.Svg));
            //ExEnd
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void AspectRatioLocked(bool isLocked)
        {
            //ExStart
            //ExFor:Shape.AspectRatioLocked
            //ExSummary:Shows how to set "AspectRatioLocked" for the shape object
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");

            //Get shape object from the document and set AspectRatioLocked(it is possible to get/set AspectRatioLocked for child shapes (mimic MS Word behavior), but AspectRatioLocked has effect only for top level shapes!)
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            shape.AspectRatioLocked = isLocked;
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.AreEqual(isLocked, shape.AspectRatioLocked);
        }

        [Test]
        public void MarkupLunguageByDefault()
        {
            //ExStart
            //ExFor:Shape.MarkupLanguage
            //ExStart:Shows how get markup language for shape object in document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape image = builder.InsertImage(MyDir + @"dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(ShapeMarkupLanguage.Dml, shape.MarkupLanguage);

                Console.WriteLine("Shape: " + shape.MarkupLanguage);
                Console.WriteLine("ShapeSize: " + shape.SizeInPoints);
            }
            //ExEnd
        }

        [Test]
        [TestCase(MsWordVersion.Word2003, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2010, ShapeMarkupLanguage.Dml)]
        public void MarkupLunguageForDifferentMsWordVersions(MsWordVersion msWordVersion, ShapeMarkupLanguage shapeMarkupLanguage)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            doc.CompatibilityOptions.OptimizeFor(msWordVersion);
            
            Shape image = builder.InsertImage(MyDir + @"dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(shapeMarkupLanguage, shape.MarkupLanguage);
            }
        }

        [Test]
        public void ChangeStrokeProperties()
        {
            //ExStart
            //ExFor:Stroke
            //ExSummary:Shows how change stroke properties
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a new shape of type Rectangle
            Shape rectangle = new Shape(doc, ShapeType.Rectangle);

            //Change stroke properties
            Stroke stroke = rectangle.Stroke;
            stroke.On = true;
            stroke.Weight = 5;
            stroke.Color = Color.Red;
            stroke.DashStyle = DashStyle.ShortDashDotDot;
            stroke.JoinStyle = JoinStyle.Miter;
            stroke.EndCap = EndCap.Square;
            stroke.LineStyle = ShapeLineStyle.Triple;

            //Insert shape object
            builder.InsertNode(rectangle);
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            rectangle = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Stroke strokeAfter = rectangle.Stroke;

            Assert.AreEqual(true, strokeAfter.On);
            Assert.AreEqual(5, strokeAfter.Weight);
            Assert.AreEqual(Color.Red.ToArgb(), strokeAfter.Color.ToArgb());
            Assert.AreEqual(DashStyle.ShortDashDotDot, strokeAfter.DashStyle);
            Assert.AreEqual(JoinStyle.Miter, strokeAfter.JoinStyle);
            Assert.AreEqual(EndCap.Square, strokeAfter.EndCap);
            Assert.AreEqual(ShapeLineStyle.Triple, strokeAfter.LineStyle);
        }
    }
}
