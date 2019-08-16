// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;
using Aspose.Words.Math;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;
using Color = System.Drawing.Color;
using DashStyle = Aspose.Words.Drawing.DashStyle;
using HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment;
using TextBox = Aspose.Words.Drawing.TextBox;
#if NETSTANDARD2_0 || __MOBILE__
using SkiaSharp;
#else
using System.Windows.Forms;
#endif

namespace ApiExamples
{
    /// <summary>
    /// Examples using shapes in documents.
    /// </summary>
    [TestFixture]
    public class ExShape : ApiExampleBase
    {
#if !(NETSTANDARD2_0 || __MOBILE__)
        [Test]
        public void InsertShape()
        {
            //ExStart
            //ExFor:ShapeBase.AlternativeText
            //ExFor:ShapeBase.Name
            //ExFor:ShapeBase.Font
            //ExFor:ShapeBase.CanHaveImage
            //ExFor:ShapeBase.ParentParagraph
            //ExFor:ShapeBase.Rotation
            //ExSummary:Shows how to insert shapes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a cube and set its name
            Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);
            shape.Name = "MyCube";
            
            // We can also set the alt text like this
            // This text will be found in Format AutoShape > Alt Text
            shape.AlternativeText = "Alt text for MyCube.";
            
            // Insert a text box
            shape = builder.InsertShape(ShapeType.TextBox, 300, 50);
            shape.Font.Name = "Times New Roman";
            
            // Move the builder into the text box and write text
            builder.MoveTo(shape.LastParagraph);
            builder.Write("Hello world!");

            // Move the builder out of the text box back into the main document
            builder.MoveTo(shape.ParentParagraph);         

            // Insert a shape with an image
            shape = builder.InsertImage(Image.FromFile(ImageDir + "Aspose.Words.gif"));
            Assert.True(shape.CanHaveImage);
            Assert.True(shape.HasImage);

            // Rotate the image
            shape.Rotation = 45.0;

            doc.Save(ArtifactsDir + "Shape.InsertShapes.docx");
            //ExEnd
        }
#endif

        [Test]
        public void ShapeCoords()
        {
            //ExStart
            //ExFor:ShapeBase.DistanceBottom
            //ExFor:ShapeBase.DistanceLeft
            //ExFor:ShapeBase.DistanceRight
            //ExFor:ShapeBase.DistanceTop
            //ExSummary:Shows how to set the wrapping distance for text that surrounds a shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle and get the text to wrap tightly around its bounds
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 150);
            shape.WrapType = WrapType.Tight;

            // Set the minimum distance between the shape and surrounding text
            shape.DistanceTop = 40.0;
            shape.DistanceBottom = 40.0;
            shape.DistanceLeft = 40.0;
            shape.DistanceRight = 40.0;

            // Move the shape closer to the centre of the page
            shape.Left = 100.0;
            shape.Top = 100.0;

            // Rotate the shape
            shape.Rotation = 60.0;

            // Add text that the shape will push out of the way
            for (int i = 0; i < 500; i++)
            {
                builder.Write("text ");
            }

            doc.Save(ArtifactsDir + "Shape.ShapeCoords.docx");
            //ExEnd
        }

        [Test]
        public void InsertGroupShape()
        {
            //ExStart
            //ExFor:ShapeBase.AnchorLocked
            //ExFor:ShapeBase.IsTopLevel
            //ExFor:ShapeBase.CoordOrigin
            //ExFor:ShapeBase.CoordSize
            //ExFor:ShapeBase.LocalToParent(PointF)
            //ExSummary:Shows how to create and work with a group of shapes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Every GroupShape is top level
            GroupShape group = new GroupShape(doc);
            Assert.True(group.IsGroup);
            Assert.True(group.IsTopLevel);

            // And it is a floating shape too, so we can set its coordinates independently of the text
            Assert.AreEqual(WrapType.None, group.WrapType);

            // Make it a floating shape
            group.WrapType = WrapType.None;

            // Top level shapes can have this property changed
            group.AnchorLocked = true;

            // Set the XY coordinates of the shape group and the size of its containing block, as it appears on the page
            group.Bounds = new RectangleF(100, 50, 200, 100);

            // Set the scale of the inner coordinates of the shape group
            // These values mean that the bottom right corner of the 200x100 outer block we set before
            // will be at x = 2000 and y = 1000, or 2000 units from the left and 1000 units from the top
            group.CoordSize = new Size(2000, 1000);

            // The coordinate origin of a shape group is x = 0, y = 0 by default, which is the top left corner
            // If we insert a child shape and set its distance from the left to 2000 and the distance from the top to 1000,
            // its origin will be at the bottom right corner of the shape group
            // We can offset the coordinate origin by setting the CoordOrigin attribute
            // In this instance, we move the origin to the centre of the shape group
            group.CoordOrigin = new Point(-1000, -500);
            
            // Populate the shape group with child shapes
            // First, insert a rectangle
            Shape subShape = new Shape(doc, ShapeType.Rectangle);
            subShape.Width = 500;
            subShape.Height = 700;

            // Place its top left corner at the parent group's coordinate origin, which is currently at its centre
            subShape.Left = 0;
            subShape.Top = 0;

            // Add the rectangle to the group
            group.AppendChild(subShape);

            // Insert a triangle
            subShape = new Shape(doc, ShapeType.Triangle);
            subShape.Width = 400;
            subShape.Height = 400;

            // Place its origin at the bottom right corner of the group
            subShape.Left = 1000;
            subShape.Top = 500;

            // The offset between this child shape and parent group can be seen here
            Assert.AreEqual(new PointF(1000, 500), subShape.LocalToParent(new PointF(0, 0)));

            // Add the triangle to the group
            group.AppendChild(subShape);

            // Child shapes of a group shape are not top level
            Assert.False(subShape.IsTopLevel);

            // Finally, insert the group into the document and save
            builder.InsertNode(group);
            doc.Save(ArtifactsDir + "Shape.InsertGroupShape.docx");
            //ExEnd
        }

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

            doc.Save(ArtifactsDir + "Shape.DeleteAllShapes.doc");
        }

        [Test]
        public void CheckShapeInline()
        {
            //ExStart
            //ExFor:ShapeBase.IsInline
            //ExSummary:Shows how to test if a shape in the document is inline or floating.
            Document doc = new Document(MyDir + "Shape.DeleteAllShapes.doc");

            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                if (shape.IsInline)
                    Console.WriteLine("Shape is inline.");
                else
                    Console.WriteLine("Shape is floating.");
            }
            //ExEnd

            // Verify that the first shape in the document is not inline.
            Assert.False(((Shape) doc.GetChild(NodeType.Shape, 0, true)).IsInline);
        }

        [Test]
        public void LineFlipOrientation()
        {
            //ExStart
            //ExFor:ShapeBase.Bounds
            //ExFor:ShapeBase.BoundsInPoints
            //ExFor:ShapeBase.FlipOrientation
            //ExFor:FlipOrientation
            //ExSummary:Shows how to create line shapes and set specific location and size.
            Document doc = new Document();

            // The lines will cross the whole page.
            float pageWidth = (float) doc.FirstSection.PageSetup.PageWidth;
            float pageHeight = (float) doc.FirstSection.PageSetup.PageHeight;

            // This line goes from top left to bottom right by default.
            Shape lineA = new Shape(doc, ShapeType.Line)
            {
                Bounds = new RectangleF(0, 0, pageWidth, pageHeight),
                RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
                RelativeVerticalPosition = RelativeVerticalPosition.Page
            };

            Assert.AreEqual(new RectangleF(0, 0, pageWidth, pageHeight), lineA.BoundsInPoints);

            // This line goes from bottom left to top right because we flipped it.
            Shape lineB = new Shape(doc, ShapeType.Line)
            {
                Bounds = new RectangleF(0, 0, pageWidth, pageHeight),
                FlipOrientation = FlipOrientation.Horizontal,
                RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
                RelativeVerticalPosition = RelativeVerticalPosition.Page
            };

            Assert.AreEqual(new RectangleF(0, 0, pageWidth, pageHeight), lineB.BoundsInPoints);

            // Add lines to the document.
            doc.FirstSection.Body.FirstParagraph.AppendChild(lineA);
            doc.FirstSection.Body.FirstParagraph.AppendChild(lineB);

            doc.Save(ArtifactsDir + "Shape.LineFlipOrientation.doc");
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

            builder.Document.Save(ArtifactsDir + "Shape.Fill.doc");
            //ExEnd
        }

        [Test]
        public void GetShapeAltTextTitle()
        {
            //ExStart
            //ExFor:ShapeBase.Title
            //ExSummary:Shows how to get or set title of shape object.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create test shape.
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.Width = 431.5;
            shape.Height = 346.35;
            shape.Title = "Alt Text Title";

            builder.InsertNode(shape);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            Console.WriteLine("Shape text: " + shape.Title);
            //ExEnd

            Assert.AreEqual("Alt Text Title", shape.Title);
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

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                // Filter out all shapes that we don't need.
                if (shape.ShapeType.Equals(ShapeType.TextBox))
                {
                    // Create a new shape that will replace the existing shape.
                    Shape image = new Shape(doc, ShapeType.Image);

                    // Load the image into the new shape.
                    image.ImageData.SetImage(ImageDir + "Hammer.wmf");

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

            doc.Save(ArtifactsDir + "Shape.ReplaceTextboxesWithImages.doc");
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
            doc.Save(ArtifactsDir + "Shape.CreateTextBox.doc");
            //ExEnd
        }

        [Test]
        public void GetActiveXControlProperties()
        {
            //ExStart
            //ExFor:OleControl
            //ExFor:Ole.OleControl.IsForms2OleControl
            //ExFor:Ole.OleControl.Name
            //ExFor:OleFormat.OleControl
            //ExFor:Forms2OleControl
            //ExFor:Forms2OleControl.Caption
            //ExFor:Forms2OleControl.Value
            //ExFor:Forms2OleControl.Enabled
            //ExFor:Forms2OleControl.Type
            //ExFor:Forms2OleControl.ChildNodes
            //ExSummary: Shows how to get ActiveX control and properties from the document.
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");

            //Get ActiveX control from the document 
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            OleControl oleControl = shape.OleFormat.OleControl;

            Assert.AreEqual(null, oleControl.Name);

            //Get ActiveX control properties
            if (oleControl.IsForms2OleControl)
            {
                Forms2OleControl checkBox = (Forms2OleControl) oleControl;
                Assert.AreEqual("Первый", checkBox.Caption);
                Assert.AreEqual("0", checkBox.Value);
                Assert.AreEqual(true, checkBox.Enabled);
                Assert.AreEqual(Forms2OleControlType.CheckBox, checkBox.Type);
                Assert.AreEqual(null, checkBox.ChildNodes);
            }
            //ExEnd
        }

        [Test]
        public void OleControl()
        {
            //ExStart
            //ExFor:OleFormat
            //ExFor:OleFormat.AutoUpdate
            //ExFor:OleFormat.IsLocked
            //ExFor:OleFormat.ProgId
            //ExFor:OleFormat.Save(Stream)
            //ExFor:OleFormat.Save(String)
            //ExFor:OleFormat.SuggestedExtension
            //ExSummary:Shows how to extract embedded OLE objects into files.
            Document doc = new Document(MyDir + "Shape.Ole.Spreadsheet.docm");

            // The first shape will contain an OLE object
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            // This object is a Microsoft Excel spreadsheet
            OleFormat oleFormat = shape.OleFormat;
            Assert.AreEqual("Excel.Sheet.12", oleFormat.ProgId);

            // Our object is neither auto updating nor locked from updates
            Assert.False(oleFormat.AutoUpdate);
            Assert.AreEqual(false, oleFormat.IsLocked);

            // If we want to extract the OLE object by saving it into our local file system, this property can tell us the relevant file extension
            Assert.AreEqual(".xlsx", oleFormat.SuggestedExtension);

            // We can save it via a stream
            using (FileStream fs = new FileStream(ArtifactsDir + "OLE spreadsheet extracted via stream" + oleFormat.SuggestedExtension, FileMode.Create))
            {
                oleFormat.Save(fs);
            }

            // We can also save it directly to a file
            oleFormat.Save(ArtifactsDir + "OLE spreadsheet saved directly" + oleFormat.SuggestedExtension);
            //ExEnd
        }

        [Test]
        public void OleLinked()
        {
            //ExStart
            //ExFor:OleFormat.IconCaption
            //ExFor:OleFormat.GetOleEntry(String)
            //ExFor:OleFormat.IsLink
            //ExFor:OleFormat.OleIcon
            //ExFor:OleFormat.SourceFullName
            //ExFor:OleFormat.SourceItem
            //ExSummary:Shows how to insert linked and unlinked OLE objects.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Embed a Microsoft Visio drawing as an OLE object into the document
            builder.InsertOleObject(ImageDir + "visio2010.vsd", "Package", false, false, null);

            // Insert a link to the file in the local file system and display it as an icon
            builder.InsertOleObject(ImageDir + "visio2010.vsd", "Package", true, true, null);
            
            // Both the OLE objects are stored within shapes
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, shapes.Count);

            // If the shape is an OLE object, it will have a valid OleFormat property
            // We can use it check if it is linked or displayed as an icon, among other things
            OleFormat oleFormat = shapes[0].OleFormat;
            Assert.AreEqual(false, oleFormat.IsLink);
            Assert.AreEqual(false, oleFormat.OleIcon);

            oleFormat = shapes[1].OleFormat;
            Assert.AreEqual(true, oleFormat.IsLink);
            Assert.AreEqual(true, oleFormat.OleIcon);

            // Get the name or the source file and verify that the whole file is linked
            Assert.True(oleFormat.SourceFullName.EndsWith(@"Images\visio2010.vsd"));
            Assert.AreEqual("", oleFormat.SourceItem);

            Assert.AreEqual("Packager", oleFormat.IconCaption);

            doc.Save(ArtifactsDir + "Shape.OleLinks.docx");

            // We can get a stream with the OLE data entry, if the object has this
            using (MemoryStream stream = oleFormat.GetOleEntry("\x0001CompObj"))
            {
                byte[] oleEntryBytes = stream.ToArray();
                Assert.AreEqual(76, oleEntryBytes.Length);
            }
            //ExEnd
        }

        [Test]
        public void OleControlCollection()
        {
            //ExStart
            //ExFor:OleFormat.Clsid
            //ExFor:Ole.Forms2OleControlCollection
            //ExFor:Ole.Forms2OleControlCollection.Count
            //ExFor:Ole.Forms2OleControlCollection.Item(Int32)
            //ExFor:Ole.NamespaceDoc
            //ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
            // Open a document that contains a Microsoft Forms OLE control with child controls
            Document doc = new Document(MyDir + "Shape.Ole.ControlCollection.docm");

            // Get the shape that contains the control
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual("6e182020-f460-11ce-9bcd-00aa00608e01", shape.OleFormat.Clsid.ToString());

            Forms2OleControl oleControl = (Forms2OleControl)shape.OleFormat.OleControl;

            // Some controls contain child controls
            Forms2OleControlCollection oleControlCollection = oleControl.ChildNodes;

            // In this case, the child controls are 3 option buttons
            Assert.AreEqual(3, oleControlCollection.Count);

            Assert.AreEqual("C#", oleControlCollection[0].Caption);
            Assert.AreEqual("1", oleControlCollection[0].Value);

            Assert.AreEqual("Visual Basic", oleControlCollection[1].Caption);
            Assert.AreEqual("0", oleControlCollection[1].Value);

            Assert.AreEqual("Delphi", oleControlCollection[2].Caption);
            Assert.AreEqual("0", oleControlCollection[2].Value);
            //ExEnd
        }

        [Test]
        public void SuggestedFileName()
        {
            //ExStart
            //ExFor:OleFormat.SuggestedFileName
            //ExSummary:Shows how to get suggested file name from the object.
            Document doc = new Document(MyDir + "Shape.SuggestedFileName.rtf");

            // Gets the file name suggested for the current embedded object if you want to save it into a file
            Shape oleShape = (Shape) doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true);
            String suggestedFileName = oleShape.OleFormat.SuggestedFileName;
            //ExEnd

            Assert.AreEqual("CSV.csv", suggestedFileName);
        }

        [Test]
        public void ObjectDidNotHaveSuggestedFileName()
        {
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");

            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            Assert.That(shape.OleFormat.SuggestedFileName, Is.Empty);
        }

        [Test]
        public void ResolutionDefaultValues()
        {
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            Assert.AreEqual(96, imageOptions.HorizontalResolution);
            Assert.AreEqual(96, imageOptions.VerticalResolution);
        }

        //For assert result of the test you need to open "Shape.OfficeMath.svg" and check that OfficeMath node is there
        [Test]
        public void SaveShapeObjectAsImage()
        {
            //ExStart
            //ExFor:OfficeMath.GetMathRenderer
            //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
            //ExSummary:Shows how to convert specific object into image
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            //Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
            OfficeMath math = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            math.GetMathRenderer().Save(ArtifactsDir + "Shape.OfficeMath.svg", new ImageSaveOptions(SaveFormat.Svg));
            //ExEnd
        }

        [Test]
        public void OfficeMathDisplayException()
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            Assert.That(() => officeMath.Justification = OfficeMathJustification.Inline,
                Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void OfficeMathDefaultValue()
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);

            Assert.AreEqual(OfficeMathDisplayType.Display, officeMath.DisplayType);
            Assert.AreEqual(OfficeMathJustification.Center, officeMath.Justification);
        }

        [Test]
        public void OfficeMathDisplayGold()
        {
            //ExStart
            //ExFor:OfficeMath
            //ExFor:OfficeMath.DisplayType
            //ExFor:OfficeMath.EquationXmlEncoding
            //ExFor:OfficeMath.Justification
            //ExFor:OfficeMath.NodeType
            //ExFor:OfficeMath.ParentParagraph
            //ExFor:OfficeMathDisplayType
            //ExFor:OfficeMathJustification
            //ExSummary:Shows how to set office math display formatting.
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);

            // OfficeMath nodes that are children of other OfficeMath nodes are always inline
            // The node we are working with is a base node, so its location and display type can be changed
            Assert.AreEqual(MathObjectType.OMathPara, officeMath.MathObjectType);
            Assert.AreEqual(NodeType.OfficeMath, officeMath.NodeType);
            Assert.AreEqual(officeMath.ParentNode, officeMath.ParentParagraph);

            // Used by OOXML and WML formats
            Assert.IsNull(officeMath.EquationXmlEncoding);

            // We can change the location and display type of the OfficeMath node
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;

            doc.Save(ArtifactsDir + "Shape.OfficeMath.docx");
            //ExEnd
            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "Shape.OfficeMath.docx", GoldsDir + "Shape.OfficeMath Gold.docx"));
        }

        [Test]
        public void CannotBeSetDisplayWithInlineJustification()
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Inline);
        }

        [Test]
        public void CannotBeSetInlineDisplayWithJustification()
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Inline;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Center);
        }

        [Test]
        public void OfficeMathDisplayNestedObjects()
        {
            Document doc = new Document(MyDir + "Shape.NestedOfficeMath.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, 0, true);

            //Always inline
            Assert.AreEqual(OfficeMathDisplayType.Inline, officeMath.DisplayType);
            Assert.AreEqual(OfficeMathJustification.Inline, officeMath.Justification);
        }

        [Test]
        [TestCase(0, MathObjectType.OMathPara)]
        [TestCase(1, MathObjectType.OMath)]
        [TestCase(2, MathObjectType.Supercript)]
        [TestCase(3, MathObjectType.Argument)]
        [TestCase(4, MathObjectType.SuperscriptPart)]
        public void WorkWithMathObjectType(int index, MathObjectType objectType)
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath) doc.GetChild(NodeType.OfficeMath, index, true);
            Assert.AreEqual(objectType, officeMath.MathObjectType);
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void AspectRatioLocked(bool isLocked)
        {
            //ExStart
            //ExFor:ShapeBase.AspectRatioLocked
            //ExSummary:Shows how to set "AspectRatioLocked" for the shape object
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");

            // Get shape object from the document and set AspectRatioLocked(it is possible to get/set AspectRatioLocked for child shapes (mimic MS Word behavior), 
            // but AspectRatioLocked has effect only for top level shapes!)
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            shape.AspectRatioLocked = isLocked;
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            Assert.AreEqual(isLocked, shape.AspectRatioLocked);
        }

        [Test]
        public void AspectRatioLockedDefaultValue()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The best place for the watermark image is in the header or footer so it is shown on every page.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
#if NETSTANDARD2_0 || __MOBILE__
            using (SKManagedStream stream = new SKManagedStream(File.OpenRead(ImageDir + "Watermark.png")))
            {
                using (SKBitmap bitmap = SKBitmap.Decode(stream))
                {
                    // Insert a floating picture.
                    Shape shape = builder.InsertImage(bitmap);
                    shape.WrapType = WrapType.None;
                    shape.BehindText = true;

                    shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                    shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

                    // Calculate image left and top position so it appears in the centre of the page.
                    shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
                    shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;

                    MemoryStream dstStream = new MemoryStream();
                    doc.Save(dstStream, SaveFormat.Docx);

                    shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
                    Assert.AreEqual(true, shape.AspectRatioLocked);
                }
            }
#else
            Image image = Image.FromFile(ImageDir + "Watermark.png");

            // Insert a floating picture.
            Shape shape = builder.InsertImage(image);
            shape.WrapType = WrapType.None;
            shape.BehindText = true;

            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Calculate image left and top position so it appears in the centre of the page.
            shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
            shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            Assert.AreEqual(true, shape.AspectRatioLocked);
#endif
        }

        [Test]
        public void MarkupLunguageByDefault()
        {
            //ExStart
            //ExFor:ShapeBase.MarkupLanguage
            //ExFor:ShapeBase.SizeInPoints
            //ExSummary:Shows how get markup language for shape object in document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(ImageDir + "dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                Assert.AreEqual(ShapeMarkupLanguage.Dml, shape.MarkupLanguage); //ExSkip

                Console.WriteLine("Shape: " + shape.MarkupLanguage);
                Console.WriteLine("ShapeSize: " + shape.SizeInPoints);
            }
            //ExEnd
        }

        [Test]
        [TestCase(MsWordVersion.Word2000, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2002, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2003, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2007, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2010, ShapeMarkupLanguage.Dml)]
        [TestCase(MsWordVersion.Word2013, ShapeMarkupLanguage.Dml)]
        [TestCase(MsWordVersion.Word2016, ShapeMarkupLanguage.Dml)]
        public void MarkupLunguageForDifferentMsWordVersions(MsWordVersion msWordVersion,
            ShapeMarkupLanguage shapeMarkupLanguage)
        {
            Document doc = new Document();
            doc.CompatibilityOptions.OptimizeFor(msWordVersion);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(ImageDir + "dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                Assert.AreEqual(shapeMarkupLanguage, shape.MarkupLanguage);
            }
        }

        [Test]
        public void ChangeStrokeProperties()
        {
            //ExStart
            //ExFor:Stroke
            //ExFor:Stroke.On
            //ExFor:Stroke.Weight
            //ExFor:Stroke.JoinStyle
            //ExFor:Stroke.LineStyle
            //ExFor:ShapeLineStyle
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

            rectangle = (Shape) doc.GetChild(NodeType.Shape, 0, true);

            Stroke strokeAfter = rectangle.Stroke;

            Assert.AreEqual(true, strokeAfter.On);
            Assert.AreEqual(5, strokeAfter.Weight);
            Assert.AreEqual(Color.Red.ToArgb(), strokeAfter.Color.ToArgb());
            Assert.AreEqual(DashStyle.ShortDashDotDot, strokeAfter.DashStyle);
            Assert.AreEqual(JoinStyle.Miter, strokeAfter.JoinStyle);
            Assert.AreEqual(EndCap.Square, strokeAfter.EndCap);
            Assert.AreEqual(ShapeLineStyle.Triple, strokeAfter.LineStyle);
        }

        [Test]
        [Description("WORDSNET-16067")]
        public void InsertOleObjectAsHtmlFile()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, false, null);

            doc.Save(ArtifactsDir + "Document.InsertedOleObject.docx");
        }

        [Test]
        [Description("WORDSNET-16085")]
        public void InsertOlePackage()
        {
            //ExStart
            //ExFor:OlePackage
            //ExFor:OleFormat.OlePackage
            //ExFor:OlePackage.FileName
            //ExFor:OlePackage.DisplayName
            //ExSummary:Shows how insert ole object as ole package and set it file name and display name.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            byte[] zipFileBytes = File.ReadAllBytes(DatabaseDir + "cat001.zip");

            using (MemoryStream stream = new MemoryStream(zipFileBytes))
            {
                Shape shape = builder.InsertOleObject(stream, "Package", true, null);

                OlePackage setOlePackage = shape.OleFormat.OlePackage;
                setOlePackage.FileName = "Cat FileName.zip";
                setOlePackage.DisplayName = "Cat DisplayName.zip";

                doc.Save(ArtifactsDir + "Shape.InsertOlePackage.docx");
            }
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.InsertOlePackage.docx");

            Shape getShape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            OlePackage getOlePackage = getShape.OleFormat.OlePackage;

            Assert.AreEqual("Cat FileName.zip", getOlePackage.FileName);
            Assert.AreEqual("Cat DisplayName.zip", getOlePackage.DisplayName);
        }

        [Test]
        public void GetAccessToOlePackage()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape oleObject = builder.InsertOleObject(MyDir + "Document.Spreadsheet.xlsx", false, false, null);
            Shape oleObjectAsOlePackage =
                builder.InsertOleObject(MyDir + "Document.Spreadsheet.xlsx", "Excel.Sheet", false, false, null);

            Assert.AreEqual(null, oleObject.OleFormat.OlePackage);
            Assert.AreEqual(typeof(OlePackage), oleObjectAsOlePackage.OleFormat.OlePackage.GetType());
        }

        [Test]
        public void ReplaceRelativeSizeToAbsolute()
        {
            Document doc = new Document(MyDir + "Shape.ShapeSize.docx");

            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);

            // Change shape size and rotation
            shape.Height = 300;
            shape.Width = 500;
            shape.Rotation = 30;

            doc.Save(ArtifactsDir + "Shape.Resize.docx");
        }

        [Test]
        public void DisplayTheShapeIntoATableCell()
        {
            //ExStart
            //ExFor:ShapeBase.IsLayoutInCell
            //ExFor:MsWordVersion
            //ExSummary:Shows how to display the shape, inside a table or outside of it.
            Document doc = new Document(MyDir + "Shape.LayoutInCell.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
            int num = 1;

            foreach (Run run in runs.OfType<Run>())
            {
                Shape watermark = new Shape(doc, ShapeType.TextPlainText);
                watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                watermark.IsLayoutInCell =
                    true; // False - display the shape outside of table cell, True - display the shape outside of table cell

                watermark.Width = 30;
                watermark.Height = 30;
                watermark.HorizontalAlignment = HorizontalAlignment.Center;
                watermark.VerticalAlignment = VerticalAlignment.Center;

                watermark.Rotation = -40;
                watermark.Fill.Color = Color.Gainsboro;
                watermark.StrokeColor = Color.Gainsboro;

                watermark.TextPath.Text = string.Format("{0}", num);
                watermark.TextPath.FontFamily = "Arial";

                watermark.Name = string.Format("WaterMark_{0}", Guid.NewGuid());
                watermark.WrapType =
                    WrapType.None; // Property will take effect only if the WrapType property is set to something other than WrapType.Inline
                watermark.BehindText = true;

                builder.MoveTo(run);
                builder.InsertNode(watermark);

                num = num + 1;
            }

            // Behaviour of MS Word on working with shapes in table cells is changed in the last versions.
            // Adding the following line is needed to make the shape displayed in center of a page.
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

            doc.Save(ArtifactsDir + "Shape.LayoutInCell.docx");
            //ExEnd
        }

        [Test]
        public void ShapeInsertion()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
            //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
            //ExSummary:Shows how to insert DML shapes into the document using a document builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // There are two ways of shape insertion
            // These methods allow inserting DML shape into the document model
            // Document must be saved in the format, which supports DML shapes, otherwise, such nodes will be converted
            // to VML shape, while document saving

            // 1. Free-floating shape insertion
            Shape freeFloatingShape = builder.InsertShape(ShapeType.TopCornersRounded, RelativeHorizontalPosition.Page, 100, RelativeVerticalPosition.Page, 100, 50, 50, WrapType.None);
            freeFloatingShape.Rotation = 30.0;
            // 2. Inline shape insertion
            Shape inlineShape = builder.InsertShape(ShapeType.DiagonalCornersRounded, 50, 50);
            inlineShape.Rotation = 30.0;

            // If you need to create "NonPrimitive" shapes, like SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
            // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
            // please save the document with "Strict" or "Transitional" compliance which allows saving shape as DML
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
            
            doc.Save(ArtifactsDir + "RotatedShape.docx", saveOptions);
            //ExEnd
        }

        //ExStart
        //ExFor:Shape.Accept(DocumentVisitor)
        //ExFor:Shape.Chart
        //ExFor:Shape.Clone(Boolean, INodeCloningListener)
        //ExFor:Shape.ExtrusionEnabled
        //ExFor:Shape.Filled
        //ExFor:Shape.HasChart
        //ExFor:Shape.OleFormat
        //ExFor:Shape.ShadowEnabled
        //ExFor:Shape.StoryType
        //ExFor:Shape.StrokeColor
        //ExFor:Shape.Stroked
        //ExFor:Shape.StrokeWeight
        //ExSummary:Shows how to iterate over all the shapes in a document.
        [Test] //ExSkip
        public void VisitShapes()
        {
            // Open a document that contains shapes
            Document doc = new Document(MyDir + "Shape.Revisions.docx");
            
            // Create a ShapeVisitor and get the document to accept it
            ShapeVisitor shapeVisitor = new ShapeVisitor();
            doc.Accept(shapeVisitor);

            // Print all the information that the visitor has collected
            Console.WriteLine(shapeVisitor.GetText());
        }

        /// <summary>
        /// DocumentVisitor implementation that collects information about visited shapes into a StringBuilder, to be printed to the console
        /// </summary>
        private class ShapeVisitor : DocumentVisitor
        {
            public ShapeVisitor()
            {
                mShapesVisited = 0;
                mTextIndentLevel = 0;
                mStringBuilder = new StringBuilder();
            }

            /// <summary>
            /// Appends a line to the StringBuilder with one prepended tab character for each indent level 
            /// </summary>
            private void AppendLine(string text)
            {
                for (int i = 0; i < mTextIndentLevel; i++)
                {
                    mStringBuilder.Append('\t');
                }

                mStringBuilder.AppendLine(text);
            }

            /// <summary>
            /// Return all the text that the StringBuilder has accumulated
            /// </summary>
            public string GetText()
            {
                return $"Shapes visited: {mShapesVisited}\n{mStringBuilder}";
            }

            /// <summary>
            /// Called when the start of a Shape node is visited
            /// </summary>
            public override VisitorAction VisitShapeStart(Shape shape)
            {
                AppendLine($"Shape found: {shape.ShapeType}");

                mTextIndentLevel++;

                if (shape.HasChart)
                    AppendLine($"Has chart: {shape.Chart.Title.Text}");

                AppendLine($"Extrusion enabled: {shape.ExtrusionEnabled}");
                AppendLine($"Shadow enabled: {shape.ShadowEnabled}");
                AppendLine($"StoryType: {shape.StoryType}");

                if (shape.Stroked)
                {
                    Assert.AreEqual(shape.Stroke.Color, shape.StrokeColor);
                    AppendLine($"Stroke colors: {shape.Stroke.Color}, {shape.Stroke.Color2}");
                    AppendLine($"Stroke weight: {shape.StrokeWeight}");

                }

                if (shape.Filled)
                    AppendLine($"Filled: {shape.FillColor}");

                if (shape.OleFormat != null)
                    AppendLine($"Ole found of type: {shape.OleFormat.ProgId}");

                if (shape.SignatureLine != null)
                    AppendLine($"Found signature line for: {shape.SignatureLine.Signer}, {shape.SignatureLine.SignerTitle}");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the end of a Shape node is visited
            /// </summary>
            public override VisitorAction VisitShapeEnd(Shape shape)
            {
                mTextIndentLevel--;
                mShapesVisited++;
                AppendLine($"End of {shape.ShapeType}");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the start of a GroupShape node is visited
            /// </summary>
            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                AppendLine($"Shape group found: {groupShape.ShapeType}");
                mTextIndentLevel++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the end of a GroupShape node is visited
            /// </summary>
            public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
            {
                mTextIndentLevel--;
                AppendLine($"End of {groupShape.ShapeType}");

                return VisitorAction.Continue;
            }

            private int mShapesVisited;
            private int mTextIndentLevel;
            private readonly StringBuilder mStringBuilder;
        }
        //ExEnd

        [Test]
        public void SignatureLine()
        {
            //ExStart
            //ExFor:Shape.SignatureLine
            //ExFor:ShapeBase.IsSignatureLine
            //ExFor:SignatureLine
            //ExFor:SignatureLine.AllowComments
            //ExFor:SignatureLine.DefaultInstructions
            //ExFor:SignatureLine.Email
            //ExFor:SignatureLine.Instructions
            //ExFor:SignatureLine.IsSigned
            //ExFor:SignatureLine.IsValid
            //ExFor:SignatureLine.ShowDate
            //ExFor:SignatureLine.Signer
            //ExFor:SignatureLine.SignerTitle
            //ExSummary:Shows how to create a line for a signature and insert it into a document.
            // Create a blank document and its DocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The SignatureLineOptions will contain all the data that the signature line will display
            SignatureLineOptions options = new SignatureLineOptions
            {
                AllowComments = true,
                DefaultInstructions = true,
                Email = "john.doe@management.com",
                Instructions = "Please sign here",
                ShowDate = true,
                Signer = "John Doe",
                SignerTitle = "Senior Manager"
            };

            // Insert the signature line, applying our SignatureLineOptions
            // We can control where the signature line will appear on the page using a combination of left/top indents and margin-relative positions
            // Since we're placing the signature line at the bottom right of the page, we will need to use negative indents to move it into view 
            Shape shape = builder.InsertSignatureLine(options, RelativeHorizontalPosition.RightMargin, -170.0, RelativeVerticalPosition.BottomMargin, -60.0, WrapType.None);
            Assert.True(shape.IsSignatureLine);

            // The SignatureLine object is a member of the shape that contains it
            SignatureLine signatureLine = shape.SignatureLine;

            Assert.AreEqual("john.doe@management.com", signatureLine.Email);
            Assert.AreEqual("John Doe", signatureLine.Signer);
            Assert.AreEqual("Senior Manager", signatureLine.SignerTitle);
            Assert.AreEqual("Please sign here", signatureLine.Instructions);
            Assert.True(signatureLine.ShowDate);

            Assert.True(signatureLine.AllowComments);
            Assert.True(signatureLine.DefaultInstructions);

            // We will be prompted to sign it when we open the document
            Assert.False(signatureLine.IsSigned);

            // The object may be valid, but the signature itself isn't until it is signed
            Assert.False(signatureLine.IsValid);

            doc.Save(ArtifactsDir + "Drawing.SignatureLine.docx");
            //ExEnd
        }

        [Test]
        public void TextBox()
        {
            //ExStart
            //ExFor:Shape.TextBox
            //ExFor:Shape.LastParagraph
            //ExFor:TextBox
            //ExFor:TextBox.FitShapeToText
            //ExFor:TextBox.InternalMarginBottom
            //ExFor:TextBox.InternalMarginLeft
            //ExFor:TextBox.InternalMarginRight
            //ExFor:TextBox.InternalMarginTop
            //ExFor:TextBox.LayoutFlow
            //ExFor:TextBox.TextBoxWrapMode
            //ExFor:TextBoxWrapMode
            //ExSummary:Shows how to insert text boxes and arrange their text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a shape that contains a TextBox
            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 100);
            TextBox textBox = textBoxShape.TextBox;

            // Move the document builder to inside the TextBox and write text
            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Vertical text");

            // Text is displayed vertically, written top to bottom
            textBox.LayoutFlow = LayoutFlow.TopToBottomIdeographic;

            // Move the builder out of the shape and back into the main document body
            builder.MoveTo(textBoxShape.ParentParagraph);

            // Insert another TextBox
            textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 100);
            textBox = textBoxShape.TextBox;

            // Apply these values to both these members to get the parent shape to defy the dimensions we set to fit tightly around the TextBox's text
            textBox.FitShapeToText = true;
            textBox.TextBoxWrapMode = TextBoxWrapMode.None;

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text fit tightly inside textbox");

            builder.MoveTo(textBoxShape.ParentParagraph);

            textBoxShape = builder.InsertShape(ShapeType.TextBox, 100, 100);
            textBox = textBoxShape.TextBox;

            // Set margins for the textbox
            textBox.InternalMarginTop = 15;
            textBox.InternalMarginBottom = 15;
            textBox.InternalMarginLeft = 15;
            textBox.InternalMarginRight = 15;

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text placed according to textbox margins");

            doc.Save(ArtifactsDir + "Drawing.TextBox.docx");
            //ExEnd
        }

        [Test]
        public void CreateNewTextBoxAndChangeTextAnchor()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set compatibility options to correctly using of VerticalAnchor property
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 100, 100);
            // Not all formats are compatible with this one
            // For most of incompatible formats AW generated a warnings on save, so use doc.WarningCallback to check it.
            textBoxShape.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
            
            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text placed bottom");

            doc.Save(ArtifactsDir + "Shape.CreateNewTextBoxAndChangeAnchor.docx");
        }

        [Test]
        public void GetTextBoxAndChangeTextAnchor()
        {
            //ExStart
            //ExFor:TextBoxAnchor
            //ExFor:TextBox.VerticalAnchor
            //ExSummary:Shows how to change text position inside textbox shape.
            Document doc = new Document(MyDir + "Shape.GetTextBoxAndChangeAnchor.docx");
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            Shape textbox = (Shape) shapes[0];
            textbox.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
            
            doc.Save(ArtifactsDir + "Shape.GetTextBoxAndChangeAnchor.docx");
            //ExEnd
        }

        //ExStart
        //ExFor:Shape.TextPath
        //ExFor:ShapeBase.IsWordArt
        //ExFor:TextPath
        //ExFor:TextPath.Bold
        //ExFor:TextPath.FitPath
        //ExFor:TextPath.FitShape
        //ExFor:TextPath.FontFamily
        //ExFor:TextPath.Italic
        //ExFor:TextPath.Kerning
        //ExFor:TextPath.On
        //ExFor:TextPath.ReverseRows
        //ExFor:TextPath.RotateLetters
        //ExFor:TextPath.SameLetterHeights
        //ExFor:TextPath.Shadow
        //ExFor:TextPath.SmallCaps
        //ExFor:TextPath.Spacing
        //ExFor:TextPath.StrikeThrough
        //ExFor:TextPath.Text
        //ExFor:TextPath.TextPathAlignment
        //ExFor:TextPath.Trim
        //ExFor:TextPath.Underline
        //ExFor:TextPath.XScale
        //ExFor:TextPathAlignment
        //ExSummary:Shows how to work with WordArt.
        [Test] //ExSkip
        public void InsertTextPaths()
        {
            Document doc = new Document();

            // Insert a WordArt object and capture the shape that contains it in a variable
            Shape shape = AppendWordArt(doc, "Bold & Italic", "Arial", 240, 24, Color.White, Color.Black, ShapeType.TextPlainText);

            // View and verify various text formatting settings
            shape.TextPath.Bold = true;
            shape.TextPath.Italic = true;

            Assert.False(shape.TextPath.Underline);
            Assert.False(shape.TextPath.Shadow);
            Assert.False(shape.TextPath.StrikeThrough);
            Assert.False(shape.TextPath.ReverseRows);
            Assert.False(shape.TextPath.XScale);
            Assert.False(shape.TextPath.Trim);
            Assert.False(shape.TextPath.SmallCaps);

            Assert.AreEqual(36.0, shape.TextPath.Size);
            Assert.AreEqual("Bold & Italic", shape.TextPath.Text);
            Assert.AreEqual(ShapeType.TextPlainText, shape.ShapeType);

            // Toggle whether or not to display text
            shape = AppendWordArt(doc, "On set to true", "Calibri", 150, 24, Color.Yellow, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.On = true;

            shape = AppendWordArt(doc, "On set to false", "Calibri", 150, 24, Color.Yellow, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.On = false;

            // Apply kerning
            shape = AppendWordArt(doc, "Kerning: VAV", "Times New Roman", 90, 24, Color.Orange, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.Kerning = true;

            shape = AppendWordArt(doc, "No kerning: VAV", "Times New Roman", 100, 24, Color.Orange, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.Kerning = false;

            // Apply custom spacing, on a scale from 0.0 (none) to 1.0 (default)
            shape = AppendWordArt(doc, "Spacing set to 0.1", "Calibri", 120, 24, Color.BlueViolet, Color.Blue, ShapeType.TextCascadeDown);
            shape.TextPath.Spacing = 0.1;

            // Rotate letters 90 degrees to the left, text is still laid out horizontally
            shape = AppendWordArt(doc, "RotateLetters", "Calibri", 200, 36, Color.GreenYellow, Color.Green, ShapeType.TextWave);
            shape.TextPath.RotateLetters = true;

            // Set the x-height to equal the cap height
            shape = AppendWordArt(doc, "Same character height for lower and UPPER case", "Calibri", 300, 24, Color.DeepSkyBlue, Color.DodgerBlue, ShapeType.TextSlantUp);
            shape.TextPath.SameLetterHeights = true;

            // By default, the size of the text will scale to always fit the size of the containing shape, overriding the text size setting
            shape = AppendWordArt(doc, "FitShape on", "Calibri", 160, 24, Color.LightBlue, Color.Blue, ShapeType.TextPlainText);
            Assert.True(shape.TextPath.FitShape);
            shape.TextPath.Size = 24.0;

            // If we set FitShape to false, the size of the text will defy the shape bounds and always keep the size value we set below
            // We can also set TextPathAlignment to align the text
            shape = AppendWordArt(doc, "FitShape off", "Calibri", 160, 24, Color.LightBlue, Color.Blue, ShapeType.TextPlainText);
            shape.TextPath.FitShape = false;
            shape.TextPath.Size = 24.0;
            shape.TextPath.TextPathAlignment = TextPathAlignment.Right;

            doc.Save(ArtifactsDir + "Drawing.TextPath.docx");
        }

        /// <summary>
        /// Insert a new paragraph with a WordArt shape inside it 
        /// </summary>
        private Shape AppendWordArt(Document doc, string text, string textFontFamily, double shapeWidth, double shapeHeight, Color wordArtFill, Color line, ShapeType wordArtShapeType)
        {
            // Insert a new paragraph
            Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));

            // Create an inline Shape, which will serve as a container for our WordArt, and append it to the paragraph
            // The shape can only be a valid WordArt shape if the ShapeType assigned here is a WordArt-designated ShapeType
            // These types will have "WordArt object" in the description and their enumerator names will start with "Text..."
            Shape shape = new Shape(doc, wordArtShapeType);
            shape.WrapType = WrapType.Inline;
            para.AppendChild(shape);

            // Set the shape's width and height
            shape.Width = shapeWidth;
            shape.Height = shapeHeight;

            // These color settings will apply to the letters of the displayed WordArt text
            shape.FillColor = wordArtFill;
            shape.StrokeColor = line;

            // The WordArt object is accessed here, and we will set the text and font like this
            shape.TextPath.Text = text;
            shape.TextPath.FontFamily = textFontFamily;
            
            return shape;
        }
        //ExEnd

        [Test]
        public void ShapeRevision()
        {
            //ExStart
            //ExFor:ShapeBase.IsDeleteRevision
            //ExFor:ShapeBase.IsInsertRevision
            //ExSummary:Shows how to work with revision shapes.
            // Open a blank document
            Document doc = new Document();

            // Insert an inline shape without tracking revisions
            Assert.False(doc.TrackRevisions);
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.WrapType = WrapType.Inline;
            shape.Width = 100.0;
            shape.Height = 100.0;
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            // Start tracking revisions and then insert another shape
            doc.StartTrackRevisions("John Doe");

            shape = new Shape(doc, ShapeType.Sun);
            shape.WrapType = WrapType.Inline;
            shape.Width = 100.0;
            shape.Height = 100.0;
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            // Get the document's shape collection which includes just the two shapes we added
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, shapes.Count);

            // Remove the first shape
            shapes[0].Remove();

            // Because we removed that shape while changes were being tracked, the shape counts as a delete revision
            Assert.AreEqual(ShapeType.Cube, shapes[0].ShapeType);
            Assert.True(shapes[0].IsDeleteRevision);

            // And we inserted another shape while tracking changes, so that shape will count as an insert revision
            Assert.AreEqual(ShapeType.Sun, shapes[1].ShapeType);
            Assert.True(shapes[1].IsInsertRevision);
            //ExEnd
        }

        [Test]
        public void MoveRevisions()
        {
            //ExStart
            //ExFor:ShapeBase.IsMoveFromRevision
            //ExFor:ShapeBase.IsMoveToRevision
            //ExSummary:Shows how to identify move revision shapes.
            // Open a document that contains a move revision
            // A move revision is when we, while changes are tracked, cut(not copy)-and-paste or highlight and drag text from one place to another
            // If inline shapes are caught up in the text movement, they will count as move revisions as well
            // Moving a floating shape will not count as a move revision
            Document doc = new Document(MyDir + "Shape.Revisions.docx");

            // The document has one shape that was moved, but shape move revisions will have two instances of that shape
            // One will be the shape at its arrival destination and the other will be the shape at its original location
            List<Shape> nc = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, nc.Count);

            // This is the move to revision, also the shape at its arrival destination
            Assert.False(nc[0].IsMoveFromRevision);
            Assert.True(nc[0].IsMoveToRevision);

            // This is the move from revision, which is the shape at its original location
            Assert.True(nc[1].IsMoveFromRevision);
            Assert.False(nc[1].IsMoveToRevision);
            //ExEnd
        }

        [Test]
        public void AdjustWithEffects()
        {
            //ExStart
            //ExFor:ShapeBase.AdjustWithEffects(RectangleF)
            //ExFor:ShapeBase.BoundsWithEffects
            //ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
            // Open a document that contains two shapes and get its shape collection
            Document doc = new Document(MyDir + "Shape.AdjustWithEffects.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(2, shapes.Count);

            // The two shapes are identical in terms of dimensions and shape type
            Assert.AreEqual(shapes[0].Width, shapes[1].Width);
            Assert.AreEqual(shapes[0].Height, shapes[1].Height);
            Assert.AreEqual(shapes[0].ShapeType, shapes[1].ShapeType);

            // However, the first shape has no effects, while the second one has a shadow and thick outline
            Assert.AreEqual(0.0, shapes[0].StrokeWeight);
            Assert.AreEqual(20.0, shapes[1].StrokeWeight);
            Assert.False(shapes[0].ShadowEnabled);
            Assert.True(shapes[1].ShadowEnabled);

            // These effects make the size of the second shape's silhouette bigger than that of the first
            // Even though the size of the rectangle that shows up when we click on these shapes in Microsoft Word is the same,
            // the practical outer bounds of the second shape are affected by the shadow and outline and are bigger
            // We can use the AdjustWithEffects method to see exactly how much bigger they are

            // The first shape has no outline or effects
            Shape shape = shapes[0];

            // Create a RectangleF object, which represents a rectangle, which we could potentially use as the coordinates and bounds for a shape
            RectangleF rectangleF = new RectangleF(200, 200, 1000, 1000);

            // Run this method to get the size of the rectangle adjusted for all of our shape's effects
            RectangleF rectangleFOut = shape.AdjustWithEffects(rectangleF);

            // Since the shape has no border-changing effects, its boundary dimensions are unaffected
            Assert.AreEqual(200, rectangleFOut.X);
            Assert.AreEqual(200, rectangleFOut.Y);
            Assert.AreEqual(1000, rectangleFOut.Width);
            Assert.AreEqual(1000, rectangleFOut.Height);

            // The final extent of the first shape, in points
            Assert.AreEqual(0, shape.BoundsWithEffects.X);
            Assert.AreEqual(0, shape.BoundsWithEffects.Y);
            Assert.AreEqual(147, shape.BoundsWithEffects.Width);
            Assert.AreEqual(147, shape.BoundsWithEffects.Height);

            // Do the same with the second shape
            shape = shapes[1];
            rectangleF = new RectangleF(200, 200, 1000, 1000);
            rectangleFOut = shape.AdjustWithEffects(rectangleF);
            
            // The shape's x/y coordinates (top left corner location) have been pushed back by the thick outline
            Assert.AreEqual(171.5, rectangleFOut.X);
            Assert.AreEqual(167, rectangleFOut.Y);

            // The width and height were also affected by the outline and shadow
            Assert.AreEqual(1045, rectangleFOut.Width);
            Assert.AreEqual(1132, rectangleFOut.Height);

            // These values are also affected by effects
            Assert.AreEqual(-28.5, shape.BoundsWithEffects.X);
            Assert.AreEqual(-33, shape.BoundsWithEffects.Y);
            Assert.AreEqual(192, shape.BoundsWithEffects.Width);
            Assert.AreEqual(279, shape.BoundsWithEffects.Height);
            //ExEnd
        }

        [Test]
        public void RenderAllShapes()
        {
            //ExStart
            //ExFor:ShapeBase.GetShapeRenderer
            //ExFor:NodeRendererBase.Save(Stream, ImageSaveOptions)
            //ExSummary:Shows how to export shapes to files in the local file system using a shape renderer.
            // Open a document that contains shapes and get its shape collection
            Document doc = new Document(MyDir + "Shape.VarietyOfShapes.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(7, shapes.Count);

            // There are 7 shapes in the document, with one group shape with 2 child shapes
            // The child shapes will be rendered but their parent group shape will be skipped, so we will see 6 output files
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                ShapeRenderer renderer = shape.GetShapeRenderer();
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
                renderer.Save(ArtifactsDir + $"Shape.ShapeRenderer {shape.Name}.png", options);
            }
            //ExEnd
        }

        [Test]
        public void OfficeMathRenderer()
        {
            //ExStart
            //ExFor:NodeRendererBase
            //ExFor:NodeRendererBase.BoundsInPoints
            //ExFor:NodeRendererBase.GetBoundsInPixels(Single, Single)
            //ExFor:NodeRendererBase.GetBoundsInPixels(Single, Single, Single)
            //ExFor:NodeRendererBase.GetOpaqueBoundsInPixels(Single, Single)
            //ExFor:NodeRendererBase.GetOpaqueBoundsInPixels(Single, Single, Single)
            //ExFor:NodeRendererBase.GetSizeInPixels(Single, Single)
            //ExFor:NodeRendererBase.GetSizeInPixels(Single, Single, Single)
            //ExFor:NodeRendererBase.OpaqueBoundsInPoints
            //ExFor:NodeRendererBase.SizeInPoints
            //ExFor:OfficeMathRenderer
            //ExFor:OfficeMathRenderer.#ctor(Math.OfficeMath)
            //ExSummary:Shows how to measure and scale shapes.
            // Open a document that contains an OfficeMath object
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            // Create a renderer for the OfficeMath object 
            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            OfficeMathRenderer renderer = new OfficeMathRenderer(officeMath);

            // We can measure the size of the image that the OfficeMath object will create when we render it
            Assert.AreEqual(117.0f, renderer.SizeInPoints.Width, 0.1f);
            Assert.AreEqual(12.9f, renderer.SizeInPoints.Height, 0.1f);

            Assert.AreEqual(117.0f, renderer.BoundsInPoints.Width, 0.1f);
            Assert.AreEqual(12.9f, renderer.BoundsInPoints.Height, 0.1f);

            // Shapes with transparent parts may return different values here
            Assert.AreEqual(117.0f, renderer.OpaqueBoundsInPoints.Width, 0.1f);
            Assert.AreEqual(14.7f, renderer.OpaqueBoundsInPoints.Height, 0.1f);

            // Get the shape size in pixels, with linear scaling to a specific DPI
            Rectangle bounds = renderer.GetBoundsInPixels(1.0f, 96.0f);
            Assert.AreEqual(156, bounds.Width);
            Assert.AreEqual(18, bounds.Height);

            // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions
            bounds = renderer.GetBoundsInPixels(1.0f, 96.0f, 150.0f);
            Assert.AreEqual(156, bounds.Width);
            Assert.AreEqual(27, bounds.Height);

            // The opaque bounds may vary here also
            bounds = renderer.GetOpaqueBoundsInPixels(1.0f, 96.0f);
            Assert.AreEqual(156, bounds.Width);
            Assert.AreEqual(20, bounds.Height);

            bounds = renderer.GetOpaqueBoundsInPixels(1.0f, 96.0f, 150.0f);
            Assert.AreEqual(156, bounds.Width);
            Assert.AreEqual(31, bounds.Height);
            //ExEnd
        }

#if !(NETSTANDARD2_0 || __MOBILE__)
        //ExStart
        //ExFor:NodeRendererBase.RenderToScale(Graphics, Single, Single, Single)
        //ExFor:NodeRendererBase.RenderToSize(Graphics, Single, Single, Single, Single)
        //ExFor:ShapeRenderer
        //ExFor:ShapeRenderer.#ctor(ShapeBase)
        //ExSummary:Shows how to render a shape with a Graphics object.
        [Test, Category("IgnoreOnJenkins")] //ExSkip
        public void DisplayShapeForm()
        {
            // Create a new ShapeForm instance and show it as a dialog box
            ShapeForm shapeForm = new ShapeForm();
            shapeForm.ShowDialog();
        }

        /// <summary>
        /// Windows Form that renders and displays shapes from a document.
        /// </summary>
        private class ShapeForm : Form
        {
            protected override void OnPaint(PaintEventArgs e)
            {
                // Set the size of the Form canvas
                this.Size = new Size(1000, 800);

                // Open a document and get its first shape, which is a chart
                Document doc = new Document(MyDir + "Shape.VarietyOfShapes.docx");
                Shape shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

                // Create a ShapeRenderer instance and a Graphics object
                // The ShapeRenderer will render the shape that is passed during construction over the Graphics object
                // Whatever is rendered on this Graphics object will be displayed on the screen inside this form
                ShapeRenderer renderer = new ShapeRenderer(shape);
                Graphics formGraphics = CreateGraphics();

                // Call this method on the renderer to render the chart in the passed Graphics object,
                // on a specified x/y coordinate and scale
                renderer.RenderToScale(formGraphics, 0, 0, 1.5f);

                // Get another shape from the document, and render it to a specific size instead of a linear scale
                GroupShape groupShape = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);
                renderer = new ShapeRenderer(groupShape);
                renderer.RenderToSize(formGraphics, 500, 400, 100, 200);
            }
        }
        //ExEnd
    #endif
    }
}