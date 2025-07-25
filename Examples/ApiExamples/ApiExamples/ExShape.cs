﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
using Aspose.Words.Tables;
using NUnit.Framework;
using Color = System.Drawing.Color;
using DashStyle = Aspose.Words.Drawing.DashStyle;
using HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment;
using TextBox = Aspose.Words.Drawing.TextBox;
using Aspose.Words.Themes;

namespace ApiExamples
{
    /// <summary>
    /// Examples using shapes in documents.
    /// </summary>
    [TestFixture]
    public class ExShape : ApiExampleBase
    {
        [Test]
        public void AltText()
        {
            //ExStart
            //ExFor:ShapeBase.AlternativeText
            //ExFor:ShapeBase.Name
            //ExSummary:Shows how to use a shape's alternative text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);
            shape.Name = "MyCube";

            shape.AlternativeText = "Alt text for MyCube.";

            // We can access the alternative text of a shape by right-clicking it, and then via "Format AutoShape" -> "Alt Text".
            doc.Save(ArtifactsDir + "Shape.AltText.docx");

            // Save the document to HTML, and then delete the linked image that belongs to our shape.
            // The browser that is reading our HTML will display the alt text in place of the missing image.
            doc.Save(ArtifactsDir + "Shape.AltText.html");
            Assert.That(File.Exists(ArtifactsDir + "Shape.AltText.001.png"), Is.True); //ExSkip
            File.Delete(ArtifactsDir + "Shape.AltText.001.png");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.AltText.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Cube, "MyCube", 150.0d, 150.0d, 0, 0, shape);
            Assert.That(shape.AlternativeText, Is.EqualTo("Alt text for MyCube."));
            Assert.That(shape.Font.Name, Is.EqualTo("Times New Roman"));

            doc = new Document(ArtifactsDir + "Shape.AltText.html");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Image, string.Empty, 151.5d, 151.5d, 0, 0, shape);
            Assert.That(shape.AlternativeText, Is.EqualTo("Alt text for MyCube."));

            TestUtil.FileContainsString(
                "<img src=\"Shape.AltText.001.png\" width=\"202\" height=\"202\" alt=\"Alt text for MyCube.\" " +
                "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />",
                ArtifactsDir + "Shape.AltText.html");
        }

        [TestCase(false)]
        [TestCase(true)]
        public void Font(bool hideShape)
        {
            //ExStart
            //ExFor:ShapeBase.Font
            //ExFor:ShapeBase.ParentParagraph
            //ExSummary:Shows how to insert a text box, and set the font of its contents.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            Shape shape = builder.InsertShape(ShapeType.TextBox, 300, 50);
            builder.MoveTo(shape.LastParagraph);
            builder.Write("This text is inside the text box.");

            // Set the "Hidden" property of the shape's "Font" object to "true" to hide the text box from sight
            // and collapse the space that it would normally occupy.
            // Set the "Hidden" property of the shape's "Font" object to "false" to leave the text box visible.
            shape.Font.Hidden = hideShape;

            // If the shape is visible, we will modify its appearance via the font object.
            if (!hideShape)
            {
                shape.Font.HighlightColor = Color.LightGray;
                shape.Font.Color = Color.Red;
                shape.Font.Underline = Underline.Dash;
            }

            // Move the builder out of the text box back into the main document.
            builder.MoveTo(shape.ParentParagraph);

            builder.Writeln("\nThis text is outside the text box.");

            doc.Save(ArtifactsDir + "Shape.Font.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Font.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.Font.Hidden, Is.EqualTo(hideShape));

            if (hideShape)
            {
                Assert.That(shape.Font.HighlightColor.ToArgb(), Is.EqualTo(Color.Empty.ToArgb()));
                Assert.That(shape.Font.Color.ToArgb(), Is.EqualTo(Color.Empty.ToArgb()));
                Assert.That(shape.Font.Underline, Is.EqualTo(Underline.None));
            }
            else
            {
                Assert.That(shape.Font.HighlightColor.ToArgb(), Is.EqualTo(Color.Silver.ToArgb()));
                Assert.That(shape.Font.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
                Assert.That(shape.Font.Underline, Is.EqualTo(Underline.Dash));
            }

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 300.0d, 50.0d, 0, 0, shape);
            Assert.That(shape.GetText().Trim(), Is.EqualTo("This text is inside the text box."));
            Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello world!\rThis text is inside the text box.\r\rThis text is outside the text box."));
        }

        [Test]
        public void Rotate()
        {
            //ExStart
            //ExFor:ShapeBase.CanHaveImage
            //ExFor:ShapeBase.Rotation
            //ExSummary:Shows how to insert and rotate an image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a shape with an image.
            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");
            Assert.That(shape.CanHaveImage, Is.True);
            Assert.That(shape.HasImage, Is.True);

            // Rotate the image 45 degrees clockwise.
            shape.Rotation = 45;

            doc.Save(ArtifactsDir + "Shape.Rotate.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Rotate.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Image, string.Empty, 300.0d, 300.0d, 0, 0, shape);
            Assert.That(shape.CanHaveImage, Is.True);
            Assert.That(shape.HasImage, Is.True);
            Assert.That(shape.Rotation, Is.EqualTo(45.0d));
        }

        [Test]
        public void Coordinates()
        {
            //ExStart
            //ExFor:ShapeBase.DistanceBottom
            //ExFor:ShapeBase.DistanceLeft
            //ExFor:ShapeBase.DistanceRight
            //ExFor:ShapeBase.DistanceTop
            //ExSummary:Shows how to set the wrapping distance for a text that surrounds a shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle and, get the text to wrap tightly around its bounds.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 150);
            shape.WrapType = WrapType.Tight;

            // Set the minimum distance between the shape and surrounding text to 40pt from all sides.
            shape.DistanceTop = 40;
            shape.DistanceBottom = 40;
            shape.DistanceLeft = 40;
            shape.DistanceRight = 40;

            // Move the shape closer to the center of the page, and then rotate the shape 60 degrees clockwise.
            shape.Top = 75;
            shape.Left = 150;
            shape.Rotation = 60;

            // Add text that will wrap around the shape.
            builder.Font.Size = 24;
            builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

            doc.Save(ArtifactsDir + "Shape.Coordinates.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Coordinates.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100002", 150.0d, 150.0d, 75.0d, 150.0d, shape);
            Assert.That(shape.DistanceBottom, Is.EqualTo(40.0d));
            Assert.That(shape.DistanceLeft, Is.EqualTo(40.0d));
            Assert.That(shape.DistanceRight, Is.EqualTo(40.0d));
            Assert.That(shape.DistanceTop, Is.EqualTo(40.0d));
            Assert.That(shape.Rotation, Is.EqualTo(60.0d));
        }

        [Test]
        public void GroupShape()
        {
            //ExStart
            //ExFor:ShapeBase.Bounds
            //ExFor:ShapeBase.CoordOrigin
            //ExFor:ShapeBase.CoordSize
            //ExSummary:Shows how to create and populate a group shape.
            Document doc = new Document();

            // Create a group shape. A group shape can display a collection of child shape nodes.
            // In Microsoft Word, clicking within the group shape's boundary or on one of the group shape's child shapes will
            // select all the other child shapes within this group and allow us to scale and move all the shapes at once.
            GroupShape group = new GroupShape(doc);

            Assert.That(group.WrapType, Is.EqualTo(WrapType.None));

            // Create a 400pt x 400pt group shape and place it at the document's floating shape coordinate origin.
            group.Bounds = new RectangleF(0, 0, 400, 400);

            // Set the group's internal coordinate plane size to 500 x 500pt. 
            // The top left corner of the group will have an x and y coordinate of (0, 0),
            // and the bottom right corner will have an x and y coordinate of (500, 500).
            group.CoordSize = new Size(500, 500);

            // Set the coordinates of the top left corner of the group to (-250, -250). 
            // The group's center will now have an x and y coordinate value of (0, 0),
            // and the bottom right corner will be at (250, 250).
            group.CoordOrigin = new Point(-250, -250);

            // Create a rectangle that will display the boundary of this group shape and add it to the group.
            Shape child1 = new Shape(doc, ShapeType.Rectangle)
            {
                Width = group.CoordSize.Width,
                Height = group.CoordSize.Height,
                Left = group.CoordOrigin.X,
                Top = group.CoordOrigin.Y
            };
            group.AppendChild(child1);

            // Once a shape is a part of a group shape, we can access it as a child node and then modify it.
            ((Shape)group.GetChild(NodeType.Shape, 0, true)).Stroke.DashStyle = DashStyle.Dash;

            // Create a small red star and insert it into the group.
            // Line up the shape with the group's coordinate origin, which we have moved to the center.
            Shape child2 = new Shape(doc, ShapeType.Star)
            {
                Width = 20,
                Height = 20,
                Left = -10,
                Top = -10,
                FillColor = Color.Red
            };
            group.AppendChild(child2);

            // Insert a rectangle, and then insert a slightly smaller rectangle in the same place with an image.
            // Newer shapes that we add to the group overlap older shapes. The light blue rectangle will partially overlap the red star,
            // and then the shape with the image will overlap the light blue rectangle, using it as a frame.
            // We cannot use the "ZOrder" properties of shapes to manipulate their arrangement within a group shape.
            Shape child3 = new Shape(doc, ShapeType.Rectangle)
            {
                Width = 250,
                Height = 250,
                Left = -250,
                Top = -250,
                FillColor = Color.LightBlue
            };
            group.AppendChild(child3);

            Shape child4 = new Shape(doc, ShapeType.Image)
            {
                Width = 200,
                Height = 200,
                Left = -225,
                Top = -225
            };
            group.AppendChild(child4);

            ((Shape)group.GetChild(NodeType.Shape, 3, true)).ImageData.SetImage(ImageDir + "Logo.jpg");

            // Insert a text box into the group shape. Set the "Left" property so that the text box's right edge
            // touches the right boundary of the group shape. Set the "Top" property so that the text box sits outside
            // the boundary of the group shape, with its top size lined up along the group shape's bottom margin.
            Shape child5 = new Shape(doc, ShapeType.TextBox)
            {
                Width = 200,
                Height = 50,
                Left = group.CoordSize.Width + group.CoordOrigin.X - 200,
                Top = group.CoordSize.Height + group.CoordOrigin.Y
            };
            group.AppendChild(child5);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertNode(group);
            builder.MoveTo(((Shape)group.GetChild(NodeType.Shape, 4, true)).AppendChild(new Paragraph(doc)));
            builder.Write("Hello world!");

            doc.Save(ArtifactsDir + "Shape.GroupShape.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.GroupShape.docx");
            group = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);

            Assert.That(group.Bounds, Is.EqualTo(new RectangleF(0, 0, 400, 400)));
            Assert.That(group.CoordSize, Is.EqualTo(new Size(500, 500)));
            Assert.That(group.CoordOrigin, Is.EqualTo(new Point(-250, -250)));

            TestUtil.VerifyShape(ShapeType.Rectangle, string.Empty, 500.0d, 500.0d, -250.0d, -250.0d, (Shape)group.GetChild(NodeType.Shape, 0, true));
            TestUtil.VerifyShape(ShapeType.Star, string.Empty, 20.0d, 20.0d, -10.0d, -10.0d, (Shape)group.GetChild(NodeType.Shape, 1, true));
            TestUtil.VerifyShape(ShapeType.Rectangle, string.Empty, 250.0d, 250.0d, -250.0d, -250.0d, (Shape)group.GetChild(NodeType.Shape, 2, true));
            TestUtil.VerifyShape(ShapeType.Image, string.Empty, 200.0d, 200.0d, -225.0d, -225.0d, (Shape)group.GetChild(NodeType.Shape, 3, true));
            TestUtil.VerifyShape(ShapeType.TextBox, string.Empty, 200.0d, 50.0d, 250.0d, 50.0d, (Shape)group.GetChild(NodeType.Shape, 4, true));
        }

        [Test]
        public void IsTopLevel()
        {
            //ExStart
            //ExFor:ShapeBase.IsTopLevel
            //ExSummary:Shows how to tell whether a shape is a part of a group shape.
            Document doc = new Document();

            Shape shape = new Shape(doc, ShapeType.Rectangle);
            shape.Width = 200;
            shape.Height = 200;
            shape.WrapType = WrapType.None;

            // A shape by default is not part of any group shape, and therefore has the "IsTopLevel" property set to "true".
            Assert.That(shape.IsTopLevel, Is.True);

            GroupShape group = new GroupShape(doc);
            group.AppendChild(shape);

            // Once we assimilate a shape into a group shape, the "IsTopLevel" property changes to "false".
            Assert.That(shape.IsTopLevel, Is.False);
            //ExEnd
        }

        [Test]
        public void LocalToParent()
        {
            //ExStart
            //ExFor:ShapeBase.CoordOrigin
            //ExFor:ShapeBase.CoordSize
            //ExFor:ShapeBase.LocalToParent(PointF)
            //ExSummary:Shows how to translate the x and y coordinate location on a shape's coordinate plane to a location on the parent shape's coordinate plane.
            Document doc = new Document();

            // Insert a group shape, and place it 100 points below and to the right of
            // the document's x and Y coordinate origin point.
            GroupShape group = new GroupShape(doc);
            group.Bounds = new RectangleF(100, 100, 500, 500);

            // Use the "LocalToParent" method to determine that (0, 0) on the group's internal x and y coordinates
            // lies on (100, 100) of its parent shape's coordinate system. The group shape's parent is the document itself.
            Assert.That(group.LocalToParent(new PointF(0, 0)), Is.EqualTo(new PointF(100, 100)));

            // By default, a shape's internal coordinate plane has the top left corner at (0, 0),
            // and the bottom right corner at (1000, 1000). Due to its size, our group shape covers an area of 500pt x 500pt
            // in the document's plane. This means that a movement of 1pt on the document's coordinate plane will translate
            // to a movement of 2pts on the group shape's coordinate plane.
            Assert.That(group.LocalToParent(new PointF(100, 100)), Is.EqualTo(new PointF(150, 150)));
            Assert.That(group.LocalToParent(new PointF(200, 200)), Is.EqualTo(new PointF(200, 200)));
            Assert.That(group.LocalToParent(new PointF(300, 300)), Is.EqualTo(new PointF(250, 250)));

            // Move the group shape's x and y axis origin from the top left corner to the center.
            // This will offset the group's internal coordinates relative to the document's coordinates even further.
            group.CoordOrigin = new Point(-250, -250);

            Assert.That(group.LocalToParent(new PointF(300, 300)), Is.EqualTo(new PointF(375, 375)));

            // Changing the scale of the coordinate plane will also affect relative locations.
            group.CoordSize = new Size(500, 500);

            Assert.That(group.LocalToParent(new PointF(300, 300)), Is.EqualTo(new PointF(650, 650)));

            // If we wish to add a shape to this group while defining its location based on a location in the document,
            // we will need to first confirm a location in the group shape that will match the document's location.
            Assert.That(group.LocalToParent(new PointF(350, 350)), Is.EqualTo(new PointF(700, 700)));

            Shape shape = new Shape(doc, ShapeType.Rectangle)
            {
                Width = 100,
                Height = 100,
                Left = 700,
                Top = 700
            };

            group.AppendChild(shape);
            doc.FirstSection.Body.FirstParagraph.AppendChild(group);

            doc.Save(ArtifactsDir + "Shape.LocalToParent.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.LocalToParent.docx");
            group = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);

            Assert.That(group.Bounds, Is.EqualTo(new RectangleF(100, 100, 500, 500)));
            Assert.That(group.CoordSize, Is.EqualTo(new Size(500, 500)));
            Assert.That(group.CoordOrigin, Is.EqualTo(new Point(-250, -250)));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AnchorLocked(bool anchorLocked)
        {
            //ExStart
            //ExFor:ShapeBase.AnchorLocked
            //ExSummary:Shows how to lock or unlock a shape's paragraph anchor.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            builder.Write("Our shape will have an anchor attached to this paragraph.");
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 160);
            shape.WrapType = WrapType.None;
            builder.InsertBreak(BreakType.ParagraphBreak);

            builder.Writeln("Hello again!");

            // Set the "AnchorLocked" property to "true" to prevent the shape's anchor
            // from moving when moving the shape in Microsoft Word.
            // Set the "AnchorLocked" property to "false" to allow any movement of the shape
            // to also move its anchor to any other paragraph that the shape ends up close to.
            shape.AnchorLocked = anchorLocked;

            // If the shape does not have a visible anchor symbol to its left,
            // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
            doc.Save(ArtifactsDir + "Shape.AnchorLocked.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.AnchorLocked.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.AnchorLocked, Is.EqualTo(anchorLocked));
        }

        [Test]
        public void DeleteAllShapes()
        {
            //ExStart
            //ExFor:Shape
            //ExSummary:Shows how to delete all shapes from a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert two shapes along with a group shape with another shape inside it.
            builder.InsertShape(ShapeType.Rectangle, 400, 200);
            builder.InsertShape(ShapeType.Star, 300, 300);

            GroupShape group = new GroupShape(doc);
            group.Bounds = new RectangleF(100, 50, 200, 100);
            group.CoordOrigin = new Point(-1000, -500);

            Shape subShape = new Shape(doc, ShapeType.Cube);
            subShape.Width = 500;
            subShape.Height = 700;
            subShape.Left = 0;
            subShape.Top = 0;

            group.AppendChild(subShape);
            builder.InsertNode(group);

            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(3));
            Assert.That(doc.GetChildNodes(NodeType.GroupShape, true).Count, Is.EqualTo(1));

            // Remove all Shape nodes from the document.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            shapes.Clear();

            // All shapes are gone, but the group shape is still in the document.
            Assert.That(doc.GetChildNodes(NodeType.GroupShape, true).Count, Is.EqualTo(1));
            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(0));

            // Remove all group shapes separately.
            NodeCollection groupShapes = doc.GetChildNodes(NodeType.GroupShape, true);
            groupShapes.Clear();

            Assert.That(doc.GetChildNodes(NodeType.GroupShape, true).Count, Is.EqualTo(0));
            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(0));
            //ExEnd
        }

        [Test]
        public void IsInline()
        {
            //ExStart
            //ExFor:ShapeBase.IsInline
            //ExSummary:Shows how to determine whether a shape is inline or floating.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two wrapping types that shapes may have.
            // 1 -  Inline:
            builder.Write("Hello world! ");
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 100);
            shape.FillColor = Color.LightBlue;
            builder.Write(" Hello again.");

            // An inline shape sits inside a paragraph among other paragraph elements, such as runs of text.
            // In Microsoft Word, we may click and drag the shape to any paragraph as if it is a character.
            // If the shape is large, it will affect vertical paragraph spacing.
            // We cannot move this shape to a place with no paragraph.
            Assert.That(shape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(shape.IsInline, Is.True);

            // 2 -  Floating:
            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 200,
                RelativeVerticalPosition.TopMargin, 200, 100, 100, WrapType.None);
            shape.FillColor = Color.Orange;

            // A floating shape belongs to the paragraph that we insert it into,
            // which we can determine by an anchor symbol that appears when we click the shape.
            // If the shape does not have a visible anchor symbol to its left,
            // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
            // In Microsoft Word, we may left click and drag this shape freely to any location.
            Assert.That(shape.WrapType, Is.EqualTo(WrapType.None));
            Assert.That(shape.IsInline, Is.False);

            doc.Save(ArtifactsDir + "Shape.IsInline.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.IsInline.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100002", 100, 100, 0, 0, shape);
            Assert.That(shape.FillColor.ToArgb(), Is.EqualTo(Color.LightBlue.ToArgb()));
            Assert.That(shape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(shape.IsInline, Is.True);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100004", 100, 100, 200, 200, shape);
            Assert.That(shape.FillColor.ToArgb(), Is.EqualTo(Color.Orange.ToArgb()));
            Assert.That(shape.WrapType, Is.EqualTo(WrapType.None));
            Assert.That(shape.IsInline, Is.False);
        }

        [Test]
        public void Bounds()
        {
            //ExStart
            //ExFor:ShapeBase.Bounds
            //ExFor:ShapeBase.BoundsInPoints
            //ExSummary:Shows how to verify shape containing block boundaries.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Line, RelativeHorizontalPosition.LeftMargin, 50,
                RelativeVerticalPosition.TopMargin, 50, 100, 100, WrapType.None);
            shape.StrokeColor = Color.Orange;

            // Even though the line itself takes up little space on the document page,
            // it occupies a rectangular containing block, the size of which we can determine using the "Bounds" properties.
            Assert.That(shape.Bounds, Is.EqualTo(new RectangleF(50, 50, 100, 100)));
            Assert.That(shape.BoundsInPoints, Is.EqualTo(new RectangleF(50, 50, 100, 100)));

            // Create a group shape, and then set the size of its containing block using the "Bounds" property.
            GroupShape group = new GroupShape(doc);
            group.Bounds = new RectangleF(0, 100, 250, 250);

            Assert.That(group.BoundsInPoints, Is.EqualTo(new RectangleF(0, 100, 250, 250)));

            // Create a rectangle, verify the size of its bounding block, and then add it to the group shape.
            shape = new Shape(doc, ShapeType.Rectangle)
            {
                Width = 100,
                Height = 100,
                Left = 700,
                Top = 700
            };

            Assert.That(shape.BoundsInPoints, Is.EqualTo(new RectangleF(700, 700, 100, 100)));

            group.AppendChild(shape);

            // The group shape's coordinate plane has its origin on the top left-hand side corner of its containing block,
            // and the x and y coordinates of (1000, 1000) on the bottom right-hand side corner.
            // Our group shape is 250x250pt in size, so every 4pt on the group shape's coordinate plane
            // translates to 1pt in the document body's coordinate plane.
            // Every shape that we insert will also shrink in size by a factor of 4.
            // The change in the shape's "BoundsInPoints" property will reflect this.
            Assert.That(shape.BoundsInPoints, Is.EqualTo(new RectangleF(175, 275, 25, 25)));

            doc.FirstSection.Body.FirstParagraph.AppendChild(group);

            // Insert a shape and place it outside of the bounds of the group shape's containing block.
            shape = new Shape(doc, ShapeType.Rectangle)
            {
                Width = 100,
                Height = 100,
                Left = 1000,
                Top = 1000
            };

            group.AppendChild(shape);

            // The group shape's footprint in the document body has increased, but the containing block remains the same.
            Assert.That(group.BoundsInPoints, Is.EqualTo(new RectangleF(0, 100, 250, 250)));
            Assert.That(shape.BoundsInPoints, Is.EqualTo(new RectangleF(250, 350, 25, 25)));

            doc.Save(ArtifactsDir + "Shape.Bounds.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Bounds.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Line, "Line 100002", 100, 100, 50, 50, shape);
            Assert.That(shape.StrokeColor.ToArgb(), Is.EqualTo(Color.Orange.ToArgb()));
            Assert.That(shape.BoundsInPoints, Is.EqualTo(new RectangleF(50, 50, 100, 100)));

            group = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);

            Assert.That(group.Bounds, Is.EqualTo(new RectangleF(0, 100, 250, 250)));
            Assert.That(group.BoundsInPoints, Is.EqualTo(new RectangleF(0, 100, 250, 250)));
            Assert.That(group.CoordSize, Is.EqualTo(new Size(1000, 1000)));
            Assert.That(group.CoordOrigin, Is.EqualTo(new Point(0, 0)));

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, string.Empty, 100, 100, 700, 700, shape);
            Assert.That(shape.BoundsInPoints, Is.EqualTo(new RectangleF(175, 275, 25, 25)));

            shape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, string.Empty, 100, 100, 1000, 1000, shape);
            Assert.That(shape.BoundsInPoints, Is.EqualTo(new RectangleF(250, 350, 25, 25)));
        }

        [Test]
        public void FlipShapeOrientation()
        {
            //ExStart
            //ExFor:ShapeBase.FlipOrientation
            //ExFor:FlipOrientation
            //ExSummary:Shows how to flip a shape on an axis.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an image shape and leave its orientation in its default state.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 100,
                RelativeVerticalPosition.TopMargin, 100, 100, 100, WrapType.None);
            shape.ImageData.SetImage(ImageDir + "Logo.jpg");

            Assert.That(shape.FlipOrientation, Is.EqualTo(FlipOrientation.None));

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 250,
                RelativeVerticalPosition.TopMargin, 100, 100, 100, WrapType.None);
            shape.ImageData.SetImage(ImageDir + "Logo.jpg");

            // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the second shape on the y-axis,
            // making it into a horizontal mirror image of the first shape.
            shape.FlipOrientation = FlipOrientation.Horizontal;

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 100,
                RelativeVerticalPosition.TopMargin, 250, 100, 100, WrapType.None);
            shape.ImageData.SetImage(ImageDir + "Logo.jpg");

            // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the third shape on the x-axis,
            // making it into a vertical mirror image of the first shape.
            shape.FlipOrientation = FlipOrientation.Vertical;

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 250,
                RelativeVerticalPosition.TopMargin, 250, 100, 100, WrapType.None);
            shape.ImageData.SetImage(ImageDir + "Logo.jpg");

            // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the fourth shape on both the x and y axes,
            // making it into a horizontal and vertical mirror image of the first shape.
            shape.FlipOrientation = FlipOrientation.Both;

            doc.Save(ArtifactsDir + "Shape.FlipShapeOrientation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.FlipShapeOrientation.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100002", 100, 100, 100, 100, shape);
            Assert.That(shape.FlipOrientation, Is.EqualTo(FlipOrientation.None));

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100004", 100, 100, 100, 250, shape);
            Assert.That(shape.FlipOrientation, Is.EqualTo(FlipOrientation.Horizontal));

            shape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100006", 100, 100, 250, 100, shape);
            Assert.That(shape.FlipOrientation, Is.EqualTo(FlipOrientation.Vertical));

            shape = (Shape)doc.GetChild(NodeType.Shape, 3, true);

            TestUtil.VerifyShape(ShapeType.Rectangle, "Rectangle 100008", 100, 100, 250, 250, shape);
            Assert.That(shape.FlipOrientation, Is.EqualTo(FlipOrientation.Both));
        }

        [Test]
        public void Fill()
        {
            //ExStart
            //ExFor:ShapeBase.Fill
            //ExFor:Shape.FillColor
            //ExFor:Shape.StrokeColor
            //ExFor:Fill
            //ExFor:Fill.Opacity
            //ExSummary:Shows how to fill a shape with a solid color.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text, and then cover it with a floating shape.
            builder.Font.Size = 32;
            builder.Writeln("Hello world!");

            Shape shape = builder.InsertShape(ShapeType.CloudCallout, RelativeHorizontalPosition.LeftMargin, 25,
                RelativeVerticalPosition.TopMargin, 25, 250, 150, WrapType.None);

            // Use the "StrokeColor" property to set the color of the outline of the shape.
            shape.StrokeColor = Color.CadetBlue;

            // Use the "FillColor" property to set the color of the inside area of the shape.
            shape.FillColor = Color.LightBlue;

            // The "Opacity" property determines how transparent the color is on a 0-1 scale,
            // with 1 being fully opaque, and 0 being invisible.
            // The shape fill by default is fully opaque, so we cannot see the text that this shape is on top of.
            Assert.That(shape.Fill.Opacity, Is.EqualTo(1.0d));

            // Set the shape fill color's opacity to a lower value so that we can see the text underneath it.
            shape.Fill.Opacity = 0.3;

            doc.Save(ArtifactsDir + "Shape.Fill.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Fill.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.CloudCallout, "CloudCallout 100002", 250.0d, 150.0d, 25.0d, 25.0d, shape);
            Color colorWithOpacity = Color.FromArgb(Convert.ToInt32(255 * shape.Fill.Opacity), Color.LightBlue.R, Color.LightBlue.G, Color.LightBlue.B);
            Assert.That(shape.FillColor.ToArgb(), Is.EqualTo(colorWithOpacity.ToArgb()));
            Assert.That(shape.StrokeColor.ToArgb(), Is.EqualTo(Color.CadetBlue.ToArgb()));
            Assert.That(shape.Fill.Opacity, Is.EqualTo(0.3d).Within(0.01d));
        }

        [Test]
        public void TextureFill()
        {
            //ExStart
            //ExFor:Fill.PresetTexture
            //ExFor:Fill.TextureAlignment
            //ExFor:TextureAlignment
            //ExSummary:Shows how to fill and tiling the texture inside the shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Rectangle, 80, 80);

            // Apply texture alignment to the shape fill.
            shape.Fill.PresetTextured(PresetTexture.Canvas);
            shape.Fill.TextureAlignment = TextureAlignment.TopRight;

            // Use the compliance option to define the shape using DML if you want to get "TextureAlignment"
            // property after the document saves.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Compliance = OoxmlCompliance.Iso29500_2008_Strict };

            doc.Save(ArtifactsDir + "Shape.TextureFill.docx", saveOptions);

            doc = new Document(ArtifactsDir + "Shape.TextureFill.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.Fill.TextureAlignment, Is.EqualTo(TextureAlignment.TopRight));
            Assert.That(shape.Fill.PresetTexture, Is.EqualTo(PresetTexture.Canvas));
            //ExEnd
        }

        [Test]
        public void GradientFill()
        {
            //ExStart
            //ExFor:Fill.OneColorGradient(Color, GradientStyle, GradientVariant, Double)
            //ExFor:Fill.OneColorGradient(GradientStyle, GradientVariant, Double)
            //ExFor:Fill.TwoColorGradient(Color, Color, GradientStyle, GradientVariant)
            //ExFor:Fill.TwoColorGradient(GradientStyle, GradientVariant)
            //ExFor:Fill.BackColor
            //ExFor:Fill.GradientStyle
            //ExFor:Fill.GradientVariant
            //ExFor:Fill.GradientAngle
            //ExFor:GradientStyle
            //ExFor:GradientVariant
            //ExSummary:Shows how to fill a shape with a gradients.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Rectangle, 80, 80);
            // Apply One-color gradient fill to the shape with ForeColor of gradient fill.
            shape.Fill.OneColorGradient(Color.Red, GradientStyle.Horizontal, GradientVariant.Variant2, 0.1);

            Assert.That(shape.Fill.ForeColor.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            Assert.That(shape.Fill.GradientStyle, Is.EqualTo(GradientStyle.Horizontal));
            Assert.That(shape.Fill.GradientVariant, Is.EqualTo(GradientVariant.Variant2));
            Assert.That(shape.Fill.GradientAngle, Is.EqualTo(270));

            shape = builder.InsertShape(ShapeType.Rectangle, 80, 80);
            // Apply Two-color gradient fill to the shape.
            shape.Fill.TwoColorGradient(GradientStyle.FromCorner, GradientVariant.Variant4);
            // Change BackColor of gradient fill.
            shape.Fill.BackColor = Color.Yellow;
            // Note that changes "GradientAngle" for "GradientStyle.FromCorner/GradientStyle.FromCenter"
            // gradient fill don't get any effect, it will work only for linear gradient.
            shape.Fill.GradientAngle = 15;

            Assert.That(shape.Fill.BackColor.ToArgb(), Is.EqualTo(Color.Yellow.ToArgb()));
            Assert.That(shape.Fill.GradientStyle, Is.EqualTo(GradientStyle.FromCorner));
            Assert.That(shape.Fill.GradientVariant, Is.EqualTo(GradientVariant.Variant4));
            Assert.That(shape.Fill.GradientAngle, Is.EqualTo(0));

            // Use the compliance option to define the shape using DML if you want to get "GradientStyle",
            // "GradientVariant" and "GradientAngle" properties after the document saves.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Compliance = OoxmlCompliance.Iso29500_2008_Strict };

            doc.Save(ArtifactsDir + "Shape.GradientFill.docx", saveOptions);
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.GradientFill.docx");
            Shape firstShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(firstShape.Fill.ForeColor.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            Assert.That(firstShape.Fill.GradientStyle, Is.EqualTo(GradientStyle.Horizontal));
            Assert.That(firstShape.Fill.GradientVariant, Is.EqualTo(GradientVariant.Variant2));
            Assert.That(firstShape.Fill.GradientAngle, Is.EqualTo(270));

            Shape secondShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.That(secondShape.Fill.BackColor.ToArgb(), Is.EqualTo(Color.Yellow.ToArgb()));
            Assert.That(secondShape.Fill.GradientStyle, Is.EqualTo(GradientStyle.FromCorner));
            Assert.That(secondShape.Fill.GradientVariant, Is.EqualTo(GradientVariant.Variant4));
            Assert.That(secondShape.Fill.GradientAngle, Is.EqualTo(0));
        }

        [Test]
        public void GradientStops()
        {
            //ExStart
            //ExFor:Fill.GradientStops
            //ExFor:GradientStopCollection
            //ExFor:GradientStopCollection.Insert(Int32, GradientStop)
            //ExFor:GradientStopCollection.Add(GradientStop)
            //ExFor:GradientStopCollection.RemoveAt(Int32)
            //ExFor:GradientStopCollection.Remove(GradientStop)
            //ExFor:GradientStopCollection.Item(Int32)
            //ExFor:GradientStopCollection.Count
            //ExFor:GradientStop
            //ExFor:GradientStop.#ctor(Color, Double)
            //ExFor:GradientStop.#ctor(Color, Double, Double)
            //ExFor:GradientStop.BaseColor
            //ExFor:GradientStop.Color
            //ExFor:GradientStop.Position
            //ExFor:GradientStop.Transparency
            //ExFor:GradientStop.Remove
            //ExSummary:Shows how to add gradient stops to the gradient fill.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Rectangle, 80, 80);
            shape.Fill.TwoColorGradient(Color.Green, Color.Red, GradientStyle.Horizontal, GradientVariant.Variant2);

            // Get gradient stops collection.
            GradientStopCollection gradientStops = shape.Fill.GradientStops;

            // Change first gradient stop.
            gradientStops[0].Color = Color.Aqua;
            gradientStops[0].Position = 0.1;
            gradientStops[0].Transparency = 0.25;

            // Add new gradient stop to the end of collection.
            GradientStop gradientStop = new GradientStop(Color.Brown, 0.5);
            gradientStops.Add(gradientStop);

            // Remove gradient stop at index 1.
            gradientStops.RemoveAt(1);
            // And insert new gradient stop at the same index 1.
            gradientStops.Insert(1, new GradientStop(Color.Chocolate, 0.75, 0.3));

            // Remove last gradient stop in the collection.
            gradientStop = gradientStops[2];
            gradientStops.Remove(gradientStop);

            Assert.That(gradientStops.Count, Is.EqualTo(2));

            Assert.That(gradientStops[0].BaseColor, Is.EqualTo(Color.FromArgb(255, 0, 255, 255)));
            Assert.That(gradientStops[0].Color.ToArgb(), Is.EqualTo(Color.Aqua.ToArgb()));
            Assert.That(gradientStops[0].Position, Is.EqualTo(0.1d).Within(0.01d));
            Assert.That(gradientStops[0].Transparency, Is.EqualTo(0.25d).Within(0.01d));

            Assert.That(gradientStops[1].Color.ToArgb(), Is.EqualTo(Color.Chocolate.ToArgb()));
            Assert.That(gradientStops[1].Position, Is.EqualTo(0.75d).Within(0.01d));
            Assert.That(gradientStops[1].Transparency, Is.EqualTo(0.3d).Within(0.01d));

            // Use the compliance option to define the shape using DML
            // if you want to get "GradientStops" property after the document saves.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Compliance = OoxmlCompliance.Iso29500_2008_Strict };

            doc.Save(ArtifactsDir + "Shape.GradientStops.docx", saveOptions);
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.GradientStops.docx");

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            gradientStops = shape.Fill.GradientStops;

            Assert.That(gradientStops.Count, Is.EqualTo(2));

            Assert.That(gradientStops[0].Color.ToArgb(), Is.EqualTo(Color.Aqua.ToArgb()));
            Assert.That(gradientStops[0].Position, Is.EqualTo(0.1d).Within(0.01d));
            Assert.That(gradientStops[0].Transparency, Is.EqualTo(0.25d).Within(0.01d));

            Assert.That(gradientStops[1].Color.ToArgb(), Is.EqualTo(Color.Chocolate.ToArgb()));
            Assert.That(gradientStops[1].Position, Is.EqualTo(0.75d).Within(0.01d));
            Assert.That(gradientStops[1].Transparency, Is.EqualTo(0.3d).Within(0.01d));
        }

        [Test]
        public void FillPattern()
        {
            //ExStart
            //ExFor:PatternType
            //ExFor:Fill.Pattern
            //ExFor:Fill.Patterned(PatternType)
            //ExFor:Fill.Patterned(PatternType, Color, Color)
            //ExSummary:Shows how to set pattern for a shape.
            Document doc = new Document(MyDir + "Shape stroke pattern border.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Fill fill = shape.Fill;

            Console.WriteLine("Pattern value is: {0}", fill.Pattern);

            // There are several ways specified fill to a pattern.
            // 1 -  Apply pattern to the shape fill:
            fill.Patterned(PatternType.DiagonalBrick);

            // 2 -  Apply pattern with foreground and background colors to the shape fill:
            fill.Patterned(PatternType.DiagonalBrick, Color.Aqua, Color.Bisque);

            doc.Save(ArtifactsDir + "Shape.FillPattern.docx");
            //ExEnd
        }

        [Test]
        public void FillThemeColor()
        {
            //ExStart
            //ExFor:Fill.ForeThemeColor
            //ExFor:Fill.BackThemeColor
            //ExFor:Fill.BackTintAndShade
            //ExSummary:Shows how to set theme color for foreground/background shape color.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.RoundRectangle, 80, 80);

            Fill fill = shape.Fill;
            fill.ForeThemeColor = ThemeColor.Dark1;
            fill.BackThemeColor = ThemeColor.Background2;

            // Note: do not use "BackThemeColor" and "BackTintAndShade" for font fill.
            if (fill.BackTintAndShade == 0)
                fill.BackTintAndShade = 0.2;

            doc.Save(ArtifactsDir + "Shape.FillThemeColor.docx");
            //ExEnd
        }

        [Test]
        public void FillTintAndShade()
        {
            //ExStart
            //ExFor:Fill.ForeTintAndShade
            //ExSummary:Shows how to manage lightening and darkening foreground font color.
            Document doc = new Document(MyDir + "Big document.docx");

            Fill textFill = doc.FirstSection.Body.FirstParagraph.Runs[0].Font.Fill;
            textFill.ForeThemeColor = ThemeColor.Accent1;
            if (textFill.ForeTintAndShade == 0)
                textFill.ForeTintAndShade = 0.5;

            doc.Save(ArtifactsDir + "Shape.FillTintAndShade.docx");
            //ExEnd
        }

        [Test]
        public void Title()
        {
            //ExStart
            //ExFor:ShapeBase.Title
            //ExSummary:Shows how to set the title of a shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a shape, give it a title, and then add it to the document.
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.Width = 200;
            shape.Height = 200;
            shape.Title = "My cube";

            builder.InsertNode(shape);

            // When we save a document with a shape that has a title,
            // Aspose.Words will store that title in the shape's Alt Text.
            doc.Save(ArtifactsDir + "Shape.Title.docx");

            doc = new Document(ArtifactsDir + "Shape.Title.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.Title, Is.EqualTo(string.Empty));
            Assert.That(shape.AlternativeText, Is.EqualTo("Title: My cube"));
            //ExEnd

            TestUtil.VerifyShape(ShapeType.Cube, string.Empty, 200.0d, 200.0d, 0.0d, 0.0d, shape);
        }

        [Test]
        public void ReplaceTextboxesWithImages()
        {
            //ExStart
            //ExFor:WrapSide
            //ExFor:ShapeBase.WrapSide
            //ExFor:NodeCollection
            //ExFor:CompositeNode.InsertAfter``1(``0,Node)
            //ExFor:NodeCollection.ToArray
            //ExSummary:Shows how to replace all textbox shapes with image shapes.
            Document doc = new Document(MyDir + "Textboxes in drawing canvas.docx");

            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.That(shapes.Count(s => s.ShapeType == ShapeType.TextBox), Is.EqualTo(3));
            Assert.That(shapes.Count(s => s.ShapeType == ShapeType.Image), Is.EqualTo(1));

            foreach (Shape shape in shapes)
            {
                if (shape.ShapeType == ShapeType.TextBox)
                {
                    Shape replacementShape = new Shape(doc, ShapeType.Image);
                    replacementShape.ImageData.SetImage(ImageDir + "Logo.jpg");
                    replacementShape.Left = shape.Left;
                    replacementShape.Top = shape.Top;
                    replacementShape.Width = shape.Width;
                    replacementShape.Height = shape.Height;
                    replacementShape.RelativeHorizontalPosition = shape.RelativeHorizontalPosition;
                    replacementShape.RelativeVerticalPosition = shape.RelativeVerticalPosition;
                    replacementShape.HorizontalAlignment = shape.HorizontalAlignment;
                    replacementShape.VerticalAlignment = shape.VerticalAlignment;
                    replacementShape.WrapType = shape.WrapType;
                    replacementShape.WrapSide = shape.WrapSide;

                    shape.ParentNode.InsertAfter(replacementShape, shape);
                    shape.Remove();
                }
            }

            shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.That(shapes.Count(s => s.ShapeType == ShapeType.TextBox), Is.EqualTo(0));
            Assert.That(shapes.Count(s => s.ShapeType == ShapeType.Image), Is.EqualTo(4));

            doc.Save(ArtifactsDir + "Shape.ReplaceTextboxesWithImages.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.ReplaceTextboxesWithImages.docx");
            Shape outShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(outShape.WrapSide, Is.EqualTo(WrapSide.Both));
        }

        [Test]
        public void CreateTextBox()
        {
            //ExStart
            //ExFor:Shape.#ctor(DocumentBase, ShapeType)
            //ExFor:Story.FirstParagraph
            //ExFor:Shape.FirstParagraph
            //ExFor:ShapeBase.WrapType
            //ExSummary:Shows how to create and format a text box.
            Document doc = new Document();

            // Create a floating text box.
            Shape textBox = new Shape(doc, ShapeType.TextBox);
            textBox.WrapType = WrapType.None;
            textBox.Height = 50;
            textBox.Width = 200;

            // Set the horizontal, and vertical alignment of the text inside the shape.
            textBox.HorizontalAlignment = HorizontalAlignment.Center;
            textBox.VerticalAlignment = VerticalAlignment.Top;

            // Add a paragraph to the text box and add a run of text that the text box will display.
            textBox.AppendChild(new Paragraph(doc));
            Paragraph para = textBox.FirstParagraph;
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            Run run = new Run(doc);
            run.Text = "Hello world!";
            para.AppendChild(run);

            doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

            doc.Save(ArtifactsDir + "Shape.CreateTextBox.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.CreateTextBox.docx");
            textBox = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.TextBox, string.Empty, 200.0d, 50.0d, 0.0d, 0.0d, textBox);
            Assert.That(textBox.WrapType, Is.EqualTo(WrapType.None));
            Assert.That(textBox.HorizontalAlignment, Is.EqualTo(HorizontalAlignment.Center));
            Assert.That(textBox.VerticalAlignment, Is.EqualTo(VerticalAlignment.Top));
            Assert.That(textBox.GetText().Trim(), Is.EqualTo("Hello world!"));
        }

        [Test]
        public void ZOrder()
        {
            //ExStart
            //ExFor:ShapeBase.ZOrder
            //ExSummary:Shows how to manipulate the order of shapes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert three different colored rectangles that partially overlap each other.
            // When we insert a shape that overlaps another shape, Aspose.Words places the newer shape on top of the old one.
            // The light green rectangle will overlap the light blue rectangle and partially obscure it,
            // and the light blue rectangle will obscure the orange rectangle.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 100,
                RelativeVerticalPosition.TopMargin, 100, 200, 200, WrapType.None);
            shape.FillColor = Color.Orange;

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 150,
                RelativeVerticalPosition.TopMargin, 150, 200, 200, WrapType.None);
            shape.FillColor = Color.LightBlue;

            shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 200,
                RelativeVerticalPosition.TopMargin, 200, 200, 200, WrapType.None);
            shape.FillColor = Color.LightGreen;

            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            // The "ZOrder" property of a shape determines its stacking priority among other overlapping shapes.
            // If two overlapping shapes have different "ZOrder" values,
            // Microsoft Word will place the shape with a higher value over the shape with the lower value. 
            // Set the "ZOrder" values of our shapes to place the first orange rectangle over the second light blue one
            // and the second light blue rectangle over the third light green rectangle.
            // This will reverse their original stacking order.
            shapes[0].ZOrder = 3;
            shapes[1].ZOrder = 2;
            shapes[2].ZOrder = 1;

            doc.Save(ArtifactsDir + "Shape.ZOrder.docx");
            //ExEnd
        }

        [Test]
        public void GetActiveXControlProperties()
        {
            //ExStart
            //ExFor:OleControl
            //ExFor:OleControl.IsForms2OleControl
            //ExFor:OleControl.Name
            //ExFor:OleFormat.OleControl
            //ExFor:Forms2OleControl
            //ExFor:Forms2OleControl.Caption
            //ExFor:Forms2OleControl.Value
            //ExFor:Forms2OleControl.Enabled
            //ExFor:Forms2OleControl.Type
            //ExFor:Forms2OleControl.ChildNodes
            //ExFor:Forms2OleControl.GroupName
            //ExSummary:Shows how to verify the properties of an ActiveX control.
            Document doc = new Document(MyDir + "ActiveX controls.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            OleControl oleControl = shape.OleFormat.OleControl;

            Assert.That(oleControl.Name, Is.EqualTo("CheckBox1"));

            if (oleControl.IsForms2OleControl)
            {
                Forms2OleControl checkBox = (Forms2OleControl)oleControl;
                Assert.That(checkBox.Caption, Is.EqualTo("First"));
                Assert.That(checkBox.Value, Is.EqualTo("0"));
                Assert.That(checkBox.Enabled, Is.EqualTo(true));
                Assert.That(checkBox.Type, Is.EqualTo(Forms2OleControlType.CheckBox));
                Assert.That(checkBox.ChildNodes, Is.EqualTo(null));
                Assert.That(checkBox.GroupName, Is.EqualTo(string.Empty));

                // Note, that you can't set GroupName for a Frame.
                checkBox.GroupName = "Aspose group name";
            }
            //ExEnd

            doc.Save(ArtifactsDir + "Shape.GetActiveXControlProperties.docx");
            doc = new Document(ArtifactsDir + "Shape.GetActiveXControlProperties.docx");

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Forms2OleControl forms2OleControl = (Forms2OleControl)shape.OleFormat.OleControl;

            Assert.That(forms2OleControl.GroupName, Is.EqualTo("Aspose group name"));
        }

        [Test]
        public void GetOleObjectRawData()
        {
            //ExStart
            //ExFor:OleFormat.GetRawData
            //ExSummary:Shows how to access the raw data of an embedded OLE object.
            Document doc = new Document(MyDir + "OLE objects.docx");

            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                OleFormat oleFormat = shape.OleFormat;
                if (oleFormat != null)
                {
                    Console.WriteLine($"This is {(oleFormat.IsLink ? "a linked" : "an embedded")} object");
                    byte[] oleRawData = oleFormat.GetRawData();

                    Assert.That(oleRawData.Length, Is.EqualTo(24576));
                }
            }
            //ExEnd
        }

        [Test]
        public void LinkedChartSourceFullName()
        {
            //ExStart
            //ExFor:Chart.SourceFullName
            //ExSummary:Shows how to get/set the full name of the external xls/xlsx document if the chart is linked.
            Document doc = new Document(MyDir + "Shape with linked chart.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            var sourceFullName = shape.Chart.SourceFullName;
            Assert.That(sourceFullName.Contains("Examples\\Data\\Spreadsheet.xlsx"), Is.True);
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
            Document doc = new Document(MyDir + "OLE spreadsheet.docm");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            // The OLE object in the first shape is a Microsoft Excel spreadsheet.
            OleFormat oleFormat = shape.OleFormat;

            Assert.That(oleFormat.ProgId, Is.EqualTo("Excel.Sheet.12"));

            // Our object is neither auto updating nor locked from updates.
            Assert.That(oleFormat.AutoUpdate, Is.False);
            Assert.That(oleFormat.IsLocked, Is.EqualTo(false));

            // If we plan on saving the OLE object to a file in the local file system,
            // we can use the "SuggestedExtension" property to determine which file extension to apply to the file.
            Assert.That(oleFormat.SuggestedExtension, Is.EqualTo(".xlsx"));

            // Below are two ways of saving an OLE object to a file in the local file system.
            // 1 -  Save it via a stream:
            using (FileStream fs = new FileStream(ArtifactsDir + "OLE spreadsheet extracted via stream" + oleFormat.SuggestedExtension, FileMode.Create))
            {
                oleFormat.Save(fs);
            }

            // 2 -  Save it directly to a filename:
            oleFormat.Save(ArtifactsDir + "OLE spreadsheet saved directly" + oleFormat.SuggestedExtension);
            //ExEnd

            Assert.That(new FileInfo(ArtifactsDir + "OLE spreadsheet extracted via stream.xlsx").Length < 8400, Is.True);
            Assert.That(new FileInfo(ArtifactsDir + "OLE spreadsheet saved directly.xlsx").Length < 8400, Is.True);
        }

        [Test]
        public void OleLinks()
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

            // Embed a Microsoft Visio drawing into the document as an OLE object.
            builder.InsertOleObject(ImageDir + "Microsoft Visio drawing.vsd", "Package", false, false, null);

            // Insert a link to the file in the local file system and display it as an icon.
            builder.InsertOleObject(ImageDir + "Microsoft Visio drawing.vsd", "Package", true, true, null);

            // Inserting OLE objects creates shapes that store these objects.
            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.That(shapes.Length, Is.EqualTo(2));
            Assert.That(shapes.Count(s => s.ShapeType == ShapeType.OleObject), Is.EqualTo(2));

            // If a shape contains an OLE object, it will have a valid "OleFormat" property,
            // which we can use to verify some aspects of the shape.
            OleFormat oleFormat = shapes[0].OleFormat;

            Assert.That(oleFormat.IsLink, Is.EqualTo(false));
            Assert.That(oleFormat.OleIcon, Is.EqualTo(false));

            oleFormat = shapes[1].OleFormat;

            Assert.That(oleFormat.IsLink, Is.EqualTo(true));
            Assert.That(oleFormat.OleIcon, Is.EqualTo(true));

            Assert.That(oleFormat.SourceFullName.EndsWith(@"Images" + Path.DirectorySeparatorChar + "Microsoft Visio drawing.vsd"), Is.True);
            Assert.That(oleFormat.SourceItem, Is.EqualTo(""));

            Assert.That(oleFormat.IconCaption, Is.EqualTo("Microsoft Visio drawing.vsd"));

            doc.Save(ArtifactsDir + "Shape.OleLinks.docx");

            // If the object contains OLE data, we can access it using a stream.
            using (MemoryStream stream = oleFormat.GetOleEntry("\x0001CompObj"))
            {
                byte[] oleEntryBytes = stream.ToArray();
                Assert.That(oleEntryBytes.Length, Is.EqualTo(76));
            }
            //ExEnd
        }

        [Test]
        public void OleControlCollection()
        {
            //ExStart
            //ExFor:OleFormat.Clsid
            //ExFor:Forms2OleControlCollection
            //ExFor:Forms2OleControlCollection.Count
            //ExFor:Forms2OleControlCollection.Item(Int32)
            //ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
            Document doc = new Document(MyDir + "OLE ActiveX controls.docm");

            // Shapes store and display OLE objects in the document's body.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.OleFormat.Clsid.ToString(), Is.EqualTo("6e182020-f460-11ce-9bcd-00aa00608e01"));

            Forms2OleControl oleControl = (Forms2OleControl)shape.OleFormat.OleControl;

            // Some OLE controls may contain child controls, such as the one in this document with three options buttons.
            Forms2OleControlCollection oleControlCollection = oleControl.ChildNodes;

            Assert.That(oleControlCollection.Count, Is.EqualTo(3));

            Assert.That(oleControlCollection[0].Caption, Is.EqualTo("C#"));
            Assert.That(oleControlCollection[0].Value, Is.EqualTo("1"));

            Assert.That(oleControlCollection[1].Caption, Is.EqualTo("Visual Basic"));
            Assert.That(oleControlCollection[1].Value, Is.EqualTo("0"));

            Assert.That(oleControlCollection[2].Caption, Is.EqualTo("Delphi"));
            Assert.That(oleControlCollection[2].Value, Is.EqualTo("0"));
            //ExEnd
        }

        [Test]
        public void SuggestedFileName()
        {
            //ExStart
            //ExFor:OleFormat.SuggestedFileName
            //ExSummary:Shows how to get an OLE object's suggested file name.
            Document doc = new Document(MyDir + "OLE shape.rtf");

            Shape oleShape = (Shape)doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true);

            // OLE objects can provide a suggested filename and extension,
            // which we can use when saving the object's contents into a file in the local file system.
            string suggestedFileName = oleShape.OleFormat.SuggestedFileName;

            Assert.That(suggestedFileName, Is.EqualTo("CSV.csv"));

            using (FileStream fileStream = new FileStream(ArtifactsDir + suggestedFileName, FileMode.Create))
            {
                oleShape.OleFormat.Save(fileStream);
            }
            //ExEnd
        }

        [Test]
        public void ObjectDidNotHaveSuggestedFileName()
        {
            Document doc = new Document(MyDir + "ActiveX controls.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.That(shape.OleFormat.SuggestedFileName, Is.EqualTo(string.Empty));
        }

        [Test]
        public void RenderOfficeMath()
        {
            //ExStart
            //ExFor:ImageSaveOptions.Scale
            //ExFor:OfficeMath.GetMathRenderer
            //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
            //ExSummary:Shows how to render an Office Math object into an image file in the local file system.
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath math = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

            // Create an "ImageSaveOptions" object to pass to the node renderer's "Save" method to modify
            // how it renders the OfficeMath node into an image.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);

            // Set the "Scale" property to 5 to render the object to five times its original size.
            saveOptions.Scale = 5;

            math.GetMathRenderer().Save(ArtifactsDir + "Shape.RenderOfficeMath.png", saveOptions);
            //ExEnd
#if !CPLUSPLUS
            if (IsRunningOnMono())
                TestUtil.VerifyImage(735, 128, ArtifactsDir + "Shape.RenderOfficeMath.png");
            else
#endif
                TestUtil.VerifyImage(813, 87, ArtifactsDir + "Shape.RenderOfficeMath.png");
        }

        [Test]
        public void OfficeMathDisplayException()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Inline);
        }

        [Test]
        public void OfficeMathDefaultValue()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 6, true);

            Assert.That(officeMath.DisplayType, Is.EqualTo(OfficeMathDisplayType.Inline));
            Assert.That(officeMath.Justification, Is.EqualTo(OfficeMathJustification.Inline));
        }

        [Test]
        public void OfficeMath()
        {
            //ExStart
            //ExFor:OfficeMath
            //ExFor:OfficeMath.DisplayType
            //ExFor:OfficeMath.Justification
            //ExFor:OfficeMath.NodeType
            //ExFor:OfficeMath.ParentParagraph
            //ExFor:OfficeMathDisplayType
            //ExFor:OfficeMathJustification
            //ExSummary:Shows how to set office math display formatting.
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

            // OfficeMath nodes that are children of other OfficeMath nodes are always inline.
            // The node we are working with is the base node to change its location and display type.
            Assert.That(officeMath.MathObjectType, Is.EqualTo(MathObjectType.OMathPara));
            Assert.That(officeMath.NodeType, Is.EqualTo(NodeType.OfficeMath));
            Assert.That(officeMath.ParentParagraph, Is.EqualTo(officeMath.ParentNode));

            // Change the location and display type of the OfficeMath node.
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;

            doc.Save(ArtifactsDir + "Shape.OfficeMath.docx");
            //ExEnd

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "Shape.OfficeMath.docx", GoldsDir + "Shape.OfficeMath Gold.docx"), Is.True);
        }

        [Test]
        public void CannotBeSetDisplayWithInlineJustification()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Inline);
        }

        [Test]
        public void CannotBeSetInlineDisplayWithJustification()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Inline;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Center);
        }

        [Test]
        public void OfficeMathDisplayNestedObjects()
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

            Assert.That(officeMath.DisplayType, Is.EqualTo(OfficeMathDisplayType.Display));
            Assert.That(officeMath.Justification, Is.EqualTo(OfficeMathJustification.Center));
        }

        [TestCase(0, MathObjectType.OMathPara)]
        [TestCase(1, MathObjectType.OMath)]
        [TestCase(2, MathObjectType.Supercript)]
        [TestCase(3, MathObjectType.Argument)]
        [TestCase(4, MathObjectType.SuperscriptPart)]
        public void WorkWithMathObjectType(int index, MathObjectType objectType)
        {
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, index, true);
            Assert.That(officeMath.MathObjectType, Is.EqualTo(objectType));
        }

        [TestCase(true)]
        [TestCase(false)]
        public void AspectRatio(bool lockAspectRatio)
        {
            //ExStart
            //ExFor:ShapeBase.AspectRatioLocked
            //ExSummary:Shows how to lock/unlock a shape's aspect ratio.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a shape. If we open this document in Microsoft Word, we can left click the shape to reveal
            // eight sizing handles around its perimeter, which we can click and drag to change its size.
            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");

            // Set the "AspectRatioLocked" property to "true" to preserve the shape's aspect ratio
            // when using any of the four diagonal sizing handles, which change both the image's height and width.
            // Using any orthogonal sizing handles that either change the height or width will still change the aspect ratio.
            // Set the "AspectRatioLocked" property to "false" to allow us to
            // freely change the image's aspect ratio with all sizing handles.
            shape.AspectRatioLocked = lockAspectRatio;

            doc.Save(ArtifactsDir + "Shape.AspectRatio.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.AspectRatio.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.AspectRatioLocked, Is.EqualTo(lockAspectRatio));
        }

        [Test]
        public void MarkupLanguageByDefault()
        {
            //ExStart
            //ExFor:ShapeBase.MarkupLanguage
            //ExFor:ShapeBase.SizeInPoints
            //ExSummary:Shows how to verify a shape's size and markup language.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertImage(ImageDir + "Transparent background logo.png");

            Assert.That(shape.MarkupLanguage, Is.EqualTo(ShapeMarkupLanguage.Dml));
            Assert.That(shape.SizeInPoints, Is.EqualTo(new SizeF(300, 300)));
            //ExEnd
        }

        [TestCase(MsWordVersion.Word2000, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2002, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2003, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2007, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2010, ShapeMarkupLanguage.Dml)]
        [TestCase(MsWordVersion.Word2013, ShapeMarkupLanguage.Dml)]
        [TestCase(MsWordVersion.Word2016, ShapeMarkupLanguage.Dml)]
        public void MarkupLanguageForDifferentMsWordVersions(MsWordVersion msWordVersion,
            ShapeMarkupLanguage shapeMarkupLanguage)
        {
            Document doc = new Document();
            doc.CompatibilityOptions.OptimizeFor(msWordVersion);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                Assert.That(shape.MarkupLanguage, Is.EqualTo(shapeMarkupLanguage));
            }
        }

        [Test]
        public void Stroke()
        {
            //ExStart
            //ExFor:Stroke
            //ExFor:Stroke.On
            //ExFor:Stroke.Weight
            //ExFor:Stroke.JoinStyle
            //ExFor:Stroke.LineStyle
            //ExFor:Stroke.Fill
            //ExFor:ShapeLineStyle
            //ExSummary:Shows how change stroke properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 100,
                RelativeVerticalPosition.TopMargin, 100, 200, 200, WrapType.None);

            // Basic shapes, such as the rectangle, have two visible parts.
            // 1 -  The fill, which applies to the area within the outline of the shape:
            shape.Fill.ForeColor = Color.White;

            // 2 -  The stroke, which marks the outline of the shape:
            // Modify various properties of this shape's stroke.
            Stroke stroke = shape.Stroke;
            stroke.On = true;
            stroke.Weight = 5;
            stroke.Color = Color.Red;
            stroke.DashStyle = DashStyle.ShortDashDotDot;
            stroke.JoinStyle = JoinStyle.Miter;
            stroke.EndCap = EndCap.Square;
            stroke.LineStyle = ShapeLineStyle.Triple;
            stroke.Fill.TwoColorGradient(Color.Red, Color.Blue, GradientStyle.Vertical, GradientVariant.Variant1);

            doc.Save(ArtifactsDir + "Shape.Stroke.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.Stroke.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            stroke = shape.Stroke;

            Assert.That(stroke.On, Is.EqualTo(true));
            Assert.That(stroke.Weight, Is.EqualTo(5));
            Assert.That(stroke.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            Assert.That(stroke.DashStyle, Is.EqualTo(DashStyle.ShortDashDotDot));
            Assert.That(stroke.JoinStyle, Is.EqualTo(JoinStyle.Miter));
            Assert.That(stroke.EndCap, Is.EqualTo(EndCap.Square));
            Assert.That(stroke.LineStyle, Is.EqualTo(ShapeLineStyle.Triple));
        }

        [Test, Description("WORDSNET-16067")]
        public void InsertOleObjectAsHtmlFile()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, false, null);

            doc.Save(ArtifactsDir + "Shape.InsertOleObjectAsHtmlFile.docx");
        }

        [Test, Description("WORDSNET-16085")]
        public void InsertOlePackage()
        {
            //ExStart
            //ExFor:OlePackage
            //ExFor:OleFormat.OlePackage
            //ExFor:OlePackage.FileName
            //ExFor:OlePackage.DisplayName
            //ExSummary:Shows how insert an OLE object into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // OLE objects allow us to open other files in the local file system using another installed application
            // in our operating system by double-clicking on the shape that contains the OLE object in the document body.
            // In this case, our external file will be a ZIP archive.
            byte[] zipFileBytes = File.ReadAllBytes(DatabaseDir + "cat001.zip");

            using (MemoryStream stream = new MemoryStream(zipFileBytes))
            {
                Shape shape = builder.InsertOleObject(stream, "Package", true, null);

                shape.OleFormat.OlePackage.FileName = "Package file name.zip";
                shape.OleFormat.OlePackage.DisplayName = "Package display name.zip";
            }

            doc.Save(ArtifactsDir + "Shape.InsertOlePackage.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.InsertOlePackage.docx");
            Shape getShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(getShape.OleFormat.OlePackage.FileName, Is.EqualTo("Package file name.zip"));
            Assert.That(getShape.OleFormat.OlePackage.DisplayName, Is.EqualTo("Package display name.zip"));
        }

        [Test]
        public void GetAccessToOlePackage()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape oleObject = builder.InsertOleObject(MyDir + "Spreadsheet.xlsx", false, false, null);
            Shape oleObjectAsOlePackage =
                builder.InsertOleObject(MyDir + "Spreadsheet.xlsx", "Excel.Sheet", false, false, null);

            Assert.That(oleObject.OleFormat.OlePackage, Is.EqualTo(null));
            Assert.That(oleObjectAsOlePackage.OleFormat.OlePackage.GetType(), Is.EqualTo(typeof(OlePackage)));
        }

        [Test]
        public void Resize()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 300);
            shape.Height = 300;
            shape.Width = 500;
            shape.Rotation = 30;

            doc.Save(ArtifactsDir + "Shape.Resize.docx");
        }

        [Test]
        public void Calendar()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            for (int i = 0; i < 31; i++)
            {
                if (i != 0 && i % 7 == 0) builder.EndRow();
                builder.InsertCell();
                builder.Write("Cell contents");
            }

            builder.EndTable();

            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
            int num = 1;

            foreach (Run run in runs.OfType<Run>())
            {
                Shape watermark = new Shape(doc, ShapeType.TextPlainText)
                {
                    RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
                    RelativeVerticalPosition = RelativeVerticalPosition.Page,
                    Width = 30,
                    Height = 30,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Rotation = -40
                };


                watermark.Fill.ForeColor = Color.Gainsboro;
                watermark.StrokeColor = Color.Gainsboro;

                watermark.TextPath.Text = string.Format("{0}", num);
                watermark.TextPath.FontFamily = "Arial";

                watermark.Name = $"Watermark_{num++}";

                watermark.BehindText = true;

                builder.MoveTo(run);
                builder.InsertNode(watermark);
            }

            doc.Save(ArtifactsDir + "Shape.Calendar.docx");

            doc = new Document(ArtifactsDir + "Shape.Calendar.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();

            Assert.That(shapes.Count, Is.EqualTo(31));

            foreach (Shape shape in shapes)
                TestUtil.VerifyShape(ShapeType.TextPlainText, $"Watermark_{shapes.IndexOf(shape) + 1}",
                    30.0d, 30.0d, 0.0d, 0.0d, shape);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void IsLayoutInCell(bool isLayoutInCell)
        {
            //ExStart
            //ExFor:ShapeBase.IsLayoutInCell
            //ExSummary:Shows how to determine how to display a shape in a table cell.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.EndTable();

            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.BottomPadding = 20;
            tableStyle.LeftPadding = 10;
            tableStyle.RightPadding = 10;
            tableStyle.TopPadding = 20;
            tableStyle.Borders.Color = Color.Black;
            tableStyle.Borders.LineStyle = LineStyle.Single;

            table.Style = tableStyle;

            builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

            Shape shape = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.LeftMargin, 50,
                RelativeVerticalPosition.TopMargin, 100, 100, 100, WrapType.None);

            // Set the "IsLayoutInCell" property to "true" to display the shape as an inline element inside the cell's paragraph.
            // The coordinate origin that will determine the shape's location will be the top left corner of the shape's cell.
            // If we re-size the cell, the shape will move to maintain the same position starting from the cell's top left.
            // Set the "IsLayoutInCell" property to "false" to display the shape as an independent floating shape.
            // The coordinate origin that will determine the shape's location will be the top left corner of the page,
            // and the shape will not respond to any re-sizing of its cell.
            shape.IsLayoutInCell = isLayoutInCell;

            // We can only apply the "IsLayoutInCell" property to floating shapes.
            shape.WrapType = WrapType.None;

            doc.Save(ArtifactsDir + "Shape.LayoutInTableCell.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.LayoutInTableCell.docx");
            table = doc.FirstSection.Body.Tables[0];
            shape = (Shape)table.FirstRow.FirstCell.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.IsLayoutInCell, Is.EqualTo(isLayoutInCell));
        }

        [Test]
        public void ShapeInsertion()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
            //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
            //ExFor:OoxmlCompliance
            //ExFor:OoxmlSaveOptions.Compliance
            //ExSummary:Shows how to insert DML shapes into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two wrapping types that shapes may have.
            // 1 -  Floating:
            builder.InsertShape(ShapeType.TopCornersRounded, RelativeHorizontalPosition.Page, 100,
                    RelativeVerticalPosition.Page, 100, 50, 50, WrapType.None);

            // 2 -  Inline:
            builder.InsertShape(ShapeType.DiagonalCornersRounded, 50, 50);

            // If you need to create "non-primitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
            // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
            // then save the document with "Strict" or "Transitional" compliance, which allows saving shape as DML.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

            doc.Save(ArtifactsDir + "Shape.ShapeInsertion.docx", saveOptions);
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.ShapeInsertion.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();

            TestUtil.VerifyShape(ShapeType.TopCornersRounded, "TopCornersRounded 100002", 50.0d, 50.0d, 100.0d, 100.0d, shapes[0]);
            TestUtil.VerifyShape(ShapeType.DiagonalCornersRounded, "DiagonalCornersRounded 100004", 50.0d, 50.0d, 0.0d, 0.0d, shapes[1]);
        }

        //ExStart
        //ExFor:Shape.Accept(DocumentVisitor)
        //ExFor:Shape.AcceptStart(DocumentVisitor)
        //ExFor:Shape.AcceptEnd(DocumentVisitor)
        //ExFor:Shape.Chart
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
            Document doc = new Document(MyDir + "Revision shape.docx");
            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(2)); //ExSkip

            ShapeAppearancePrinter visitor = new ShapeAppearancePrinter();
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
        }

        /// <summary>
        /// Logs appearance-related information about visited shapes.
        /// </summary>
        private class ShapeAppearancePrinter : DocumentVisitor
        {
            public ShapeAppearancePrinter()
            {
                mShapesVisited = 0;
                mTextIndentLevel = 0;
                mStringBuilder = new StringBuilder();
            }

            /// <summary>
            /// Appends a line to the StringBuilder with one prepended tab character for each indent level.
            /// </summary>
            private void AppendLine(string text)
            {
                for (int i = 0; i < mTextIndentLevel; i++) mStringBuilder.Append('\t');

                mStringBuilder.AppendLine(text);
            }

            /// <summary>
            /// Return all the text that the StringBuilder has accumulated.
            /// </summary>
            public string GetText()
            {
                return $"Shapes visited: {mShapesVisited}\n{mStringBuilder}";
            }

            /// <summary>
            /// Called when this visitor visits the start of a Shape node.
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
                    Assert.That(shape.StrokeColor, Is.EqualTo(shape.Stroke.Color));
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
            /// Called when this visitor visits the end of a Shape node.
            /// </summary>
            public override VisitorAction VisitShapeEnd(Shape shape)
            {
                mTextIndentLevel--;
                mShapesVisited++;
                AppendLine($"End of {shape.ShapeType}");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when this visitor visits the start of a GroupShape node.
            /// </summary>
            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                AppendLine($"Shape group found: {groupShape.ShapeType}");
                mTextIndentLevel++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when this visitor visits the end of a GroupShape node.
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
            //ExFor:SignatureLine.ShowDate
            //ExFor:SignatureLine.Signer
            //ExFor:SignatureLine.SignerTitle
            //ExSummary:Shows how to create a line for a signature and insert it into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

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

            // Insert a shape that will contain a signature line, whose appearance we will
            // customize using the "SignatureLineOptions" object we have created above.
            // If we insert a shape whose coordinates originate at the bottom right hand corner of the page,
            // we will need to supply negative x and y coordinates to bring the shape into view.
            Shape shape = builder.InsertSignatureLine(options, RelativeHorizontalPosition.RightMargin, -170.0,
                    RelativeVerticalPosition.BottomMargin, -60.0, WrapType.None);

            Assert.That(shape.IsSignatureLine, Is.True);

            // Verify the properties of our signature line via its Shape object.
            SignatureLine signatureLine = shape.SignatureLine;

            Assert.That(signatureLine.Email, Is.EqualTo("john.doe@management.com"));
            Assert.That(signatureLine.Signer, Is.EqualTo("John Doe"));
            Assert.That(signatureLine.SignerTitle, Is.EqualTo("Senior Manager"));
            Assert.That(signatureLine.Instructions, Is.EqualTo("Please sign here"));
            Assert.That(signatureLine.ShowDate, Is.True);
            Assert.That(signatureLine.AllowComments, Is.True);
            Assert.That(signatureLine.DefaultInstructions, Is.True);

            doc.Save(ArtifactsDir + "Shape.SignatureLine.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.SignatureLine.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.Image, string.Empty, 192.75d, 96.75d, -60.0d, -170.0d, shape);
            Assert.That(shape.IsSignatureLine, Is.True);

            signatureLine = shape.SignatureLine;

            Assert.That(signatureLine.Email, Is.EqualTo("john.doe@management.com"));
            Assert.That(signatureLine.Signer, Is.EqualTo("John Doe"));
            Assert.That(signatureLine.SignerTitle, Is.EqualTo("Senior Manager"));
            Assert.That(signatureLine.Instructions, Is.EqualTo("Please sign here"));
            Assert.That(signatureLine.ShowDate, Is.True);
            Assert.That(signatureLine.AllowComments, Is.True);
            Assert.That(signatureLine.DefaultInstructions, Is.True);
            Assert.That(signatureLine.IsSigned, Is.False);
            Assert.That(signatureLine.IsValid, Is.False);
        }

        [TestCase(LayoutFlow.Vertical)]
        [TestCase(LayoutFlow.Horizontal)]
        [TestCase(LayoutFlow.HorizontalIdeographic)]
        [TestCase(LayoutFlow.BottomToTop)]
        [TestCase(LayoutFlow.TopToBottom)]
        [TestCase(LayoutFlow.TopToBottomIdeographic)]
        public void TextBoxLayoutFlow(LayoutFlow layoutFlow)
        {
            //ExStart
            //ExFor:Shape.TextBox
            //ExFor:Shape.LastParagraph
            //ExFor:TextBox
            //ExFor:TextBox.LayoutFlow
            //ExSummary:Shows how to set the orientation of text inside a text box.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 100);
            TextBox textBox = textBoxShape.TextBox;

            // Move the document builder to inside the TextBox and add text.
            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Writeln("Hello world!");
            builder.Write("Hello again!");

            // Set the "LayoutFlow" property to set an orientation for the text contents of this text box.
            textBox.LayoutFlow = layoutFlow;

            doc.Save(ArtifactsDir + "Shape.TextBoxLayoutFlow.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.TextBoxLayoutFlow.docx");
            textBoxShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 150.0d, 100.0d, 0.0d, 0.0d, textBoxShape);

            LayoutFlow expectedLayoutFlow;

            switch (layoutFlow)
            {
                case LayoutFlow.BottomToTop:
                case LayoutFlow.Horizontal:
                case LayoutFlow.TopToBottomIdeographic:
                case LayoutFlow.Vertical:
                    expectedLayoutFlow = layoutFlow;
                    break;
                case LayoutFlow.TopToBottom:
                    expectedLayoutFlow = LayoutFlow.Vertical;
                    break;
                default:
                    expectedLayoutFlow = LayoutFlow.Horizontal;
                    break;
            }

            TestUtil.VerifyTextBox(expectedLayoutFlow, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, textBoxShape.TextBox);
            Assert.That(textBoxShape.GetText().Trim(), Is.EqualTo("Hello world!\rHello again!"));
        }

        [Test]
        public void TextBoxFitShapeToText()
        {
            //ExStart
            //ExFor:TextBox
            //ExFor:TextBox.FitShapeToText
            //ExSummary:Shows how to get a text box to resize itself to fit its contents tightly.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 100);
            TextBox textBox = textBoxShape.TextBox;

            // Apply these values to both these members to get the parent shape to fit
            // tightly around the text contents, ignoring the dimensions we have set.
            textBox.FitShapeToText = true;
            textBox.TextBoxWrapMode = TextBoxWrapMode.None;

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text fit tightly inside textbox.");

            doc.Save(ArtifactsDir + "Shape.TextBoxFitShapeToText.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.TextBoxFitShapeToText.docx");
            textBoxShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 150.0d, 100.0d, 0.0d, 0.0d, textBoxShape);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, true, TextBoxWrapMode.None, 3.6d, 3.6d, 7.2d, 7.2d, textBoxShape.TextBox);
            Assert.That(textBoxShape.GetText().Trim(), Is.EqualTo("Text fit tightly inside textbox."));
        }

        [Test]
        public void TextBoxMargins()
        {
            //ExStart
            //ExFor:TextBox
            //ExFor:TextBox.InternalMarginBottom
            //ExFor:TextBox.InternalMarginLeft
            //ExFor:TextBox.InternalMarginRight
            //ExFor:TextBox.InternalMarginTop
            //ExSummary:Shows how to set internal margins for a text box.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert another textbox with specific margins.
            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox = textBoxShape.TextBox;
            textBox.InternalMarginTop = 15;
            textBox.InternalMarginBottom = 15;
            textBox.InternalMarginLeft = 15;
            textBox.InternalMarginRight = 15;

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text placed according to textbox margins.");

            doc.Save(ArtifactsDir + "Shape.TextBoxMargins.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.TextBoxMargins.docx");
            textBoxShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 100.0d, 100.0d, 0.0d, 0.0d, textBoxShape);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 15.0d, 15.0d, 15.0d, 15.0d, textBoxShape.TextBox);
            Assert.That(textBoxShape.GetText().Trim(), Is.EqualTo("Text placed according to textbox margins."));
        }

        [TestCase(TextBoxWrapMode.None)]
        [TestCase(TextBoxWrapMode.Square)]
        public void TextBoxContentsWrapMode(TextBoxWrapMode textBoxWrapMode)
        {
            //ExStart
            //ExFor:TextBox.TextBoxWrapMode
            //ExFor:TextBoxWrapMode
            //ExSummary:Shows how to set a wrapping mode for the contents of a text box.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 300, 300);
            TextBox textBox = textBoxShape.TextBox;

            // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.None" to increase the text box's width
            // to accommodate text, should it be large enough.
            // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.Square" to
            // wrap all text inside the text box, preserving its dimensions.
            textBox.TextBoxWrapMode = textBoxWrapMode;

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Font.Size = 32;
            builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            doc.Save(ArtifactsDir + "Shape.TextBoxContentsWrapMode.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.TextBoxContentsWrapMode.docx");
            textBoxShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 300.0d, 300.0d, 0.0d, 0.0d, textBoxShape);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, textBoxWrapMode, 3.6d, 3.6d, 7.2d, 7.2d, textBoxShape.TextBox);
            Assert.That(textBoxShape.GetText().Trim(), Is.EqualTo("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."));
        }

        [Test]
        public void TextBoxShapeType()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set compatibility options to correctly using of VerticalAnchor property.
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 100, 100);
            // Not all formats are compatible with this one.
            // For most of the incompatible formats, AW generated warnings on save, so use doc.WarningCallback to check it.
            textBoxShape.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Text placed bottom");

            doc.Save(ArtifactsDir + "Shape.TextBoxShapeType.docx");
        }

        [Test]
        public void CreateLinkBetweenTextBoxes()
        {
            //ExStart
            //ExFor:TextBox.IsValidLinkTarget(TextBox)
            //ExFor:TextBox.Next
            //ExFor:TextBox.Previous
            //ExFor:TextBox.BreakForwardLink
            //ExSummary:Shows how to link text boxes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textBoxShape1 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox1 = textBoxShape1.TextBox;
            builder.Writeln();

            Shape textBoxShape2 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox2 = textBoxShape2.TextBox;
            builder.Writeln();

            Shape textBoxShape3 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox3 = textBoxShape3.TextBox;
            builder.Writeln();

            Shape textBoxShape4 = builder.InsertShape(ShapeType.TextBox, 100, 100);
            TextBox textBox4 = textBoxShape4.TextBox;

            // Create links between some of the text boxes.
            if (textBox1.IsValidLinkTarget(textBox2))
                textBox1.Next = textBox2;

            if (textBox2.IsValidLinkTarget(textBox3))
                textBox2.Next = textBox3;

            // Only an empty text box may have a link.
            Assert.That(textBox3.IsValidLinkTarget(textBox4), Is.True);

            builder.MoveTo(textBoxShape4.LastParagraph);
            builder.Write("Hello world!");

            Assert.That(textBox3.IsValidLinkTarget(textBox4), Is.False);

            if (textBox1.Next != null && textBox1.Previous == null)
                Console.WriteLine("This TextBox is the head of the sequence");

            if (textBox2.Next != null && textBox2.Previous != null)
                Console.WriteLine("This TextBox is the middle of the sequence");

            if (textBox3.Next == null && textBox3.Previous != null)
            {
                Console.WriteLine("This TextBox is the tail of the sequence");

                // Break the forward link between textBox2 and textBox3, and then verify that they are no longer linked.
                textBox3.Previous.BreakForwardLink();
                Assert.That(textBox2.Next == null, Is.True);
                Assert.That(textBox3.Previous == null, Is.True);
            }

            doc.Save(ArtifactsDir + "Shape.CreateLinkBetweenTextBoxes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.CreateLinkBetweenTextBoxes.docx");
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToList();

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 100.0d, 100.0d, 0.0d, 0.0d, shapes[0]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[0].TextBox);
            Assert.That(shapes[0].GetText().Trim(), Is.EqualTo(string.Empty));

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100004", 100.0d, 100.0d, 0.0d, 0.0d, shapes[1]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[1].TextBox);
            Assert.That(shapes[1].GetText().Trim(), Is.EqualTo(string.Empty));

            TestUtil.VerifyShape(ShapeType.Rectangle, "TextBox 100006", 100.0d, 100.0d, 0.0d, 0.0d, shapes[2]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[2].TextBox);
            Assert.That(shapes[2].GetText().Trim(), Is.EqualTo(string.Empty));

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100008", 100.0d, 100.0d, 0.0d, 0.0d, shapes[3]);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shapes[3].TextBox);
            Assert.That(shapes[3].GetText().Trim(), Is.EqualTo("Hello world!"));
        }

        [TestCase(TextBoxAnchor.Top)]
        [TestCase(TextBoxAnchor.Middle)]
        [TestCase(TextBoxAnchor.Bottom)]
        public void VerticalAnchor(TextBoxAnchor verticalAnchor)
        {
            //ExStart
            //ExFor:CompatibilityOptions
            //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
            //ExFor:TextBoxAnchor
            //ExFor:TextBox.VerticalAnchor
            //ExSummary:Shows how to vertically align the text contents of a text box.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.TextBox, 200, 200);

            // Set the "VerticalAnchor" property to "TextBoxAnchor.Top" to
            // align the text in this text box with the top side of the shape.
            // Set the "VerticalAnchor" property to "TextBoxAnchor.Middle" to
            // align the text in this text box to the center of the shape.
            // Set the "VerticalAnchor" property to "TextBoxAnchor.Bottom" to
            // align the text in this text box to the bottom of the shape.
            shape.TextBox.VerticalAnchor = verticalAnchor;

            builder.MoveTo(shape.FirstParagraph);
            builder.Write("Hello world!");

            // The vertical aligning of text inside text boxes is available from Microsoft Word 2007 onwards.
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2007);
            doc.Save(ArtifactsDir + "Shape.VerticalAnchor.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.VerticalAnchor.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyShape(ShapeType.TextBox, "TextBox 100002", 200.0d, 200.0d, 0.0d, 0.0d, shape);
            TestUtil.VerifyTextBox(LayoutFlow.Horizontal, false, TextBoxWrapMode.Square, 3.6d, 3.6d, 7.2d, 7.2d, shape.TextBox);
            Assert.That(shape.TextBox.VerticalAnchor, Is.EqualTo(verticalAnchor));
            Assert.That(shape.GetText().Trim(), Is.EqualTo("Hello world!"));
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
        //ExFor:TextPath.Size
        //ExFor:TextPathAlignment
        //ExSummary:Shows how to work with WordArt.
        [Test] //ExSkip
        public void InsertTextPaths()
        {
            Document doc = new Document();

            // Insert a WordArt object to display text in a shape that we can re-size and move by using the mouse in Microsoft Word.
            // Provide a "ShapeType" as an argument to set a shape for the WordArt.
            Shape shape = AppendWordArt(doc, "Hello World! This text is bold, and italic.",
                "Arial", 480, 24, Color.White, Color.Black, ShapeType.TextPlainText);

            // Apply the "Bold" and "Italic" formatting settings to the text using the respective properties.
            shape.TextPath.Bold = true;
            shape.TextPath.Italic = true;

            // Below are various other text formatting-related properties.
            Assert.That(shape.TextPath.Underline, Is.False);
            Assert.That(shape.TextPath.Shadow, Is.False);
            Assert.That(shape.TextPath.StrikeThrough, Is.False);
            Assert.That(shape.TextPath.ReverseRows, Is.False);
            Assert.That(shape.TextPath.XScale, Is.False);
            Assert.That(shape.TextPath.Trim, Is.False);
            Assert.That(shape.TextPath.SmallCaps, Is.False);

            Assert.That(shape.TextPath.Size, Is.EqualTo(36.0));
            Assert.That(shape.TextPath.Text, Is.EqualTo("Hello World! This text is bold, and italic."));
            Assert.That(shape.ShapeType, Is.EqualTo(ShapeType.TextPlainText));

            // Use the "On" property to show/hide the text.
            shape = AppendWordArt(doc, "On set to \"true\"", "Calibri", 150, 24, Color.Yellow, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.On = true;

            shape = AppendWordArt(doc, "On set to \"false\"", "Calibri", 150, 24, Color.Yellow, Color.Purple, ShapeType.TextPlainText);
            shape.TextPath.On = false;

            // Use the "Kerning" property to enable/disable kerning spacing between certain characters.
            shape = AppendWordArt(doc, "Kerning: VAV", "Times New Roman", 90, 24, Color.Orange, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.Kerning = true;

            shape = AppendWordArt(doc, "No kerning: VAV", "Times New Roman", 100, 24, Color.Orange, Color.Red, ShapeType.TextPlainText);
            shape.TextPath.Kerning = false;

            // Use the "Spacing" property to set the custom spacing between characters on a scale from 0.0 (none) to 1.0 (default).
            shape = AppendWordArt(doc, "Spacing set to 0.1", "Calibri", 120, 24, Color.BlueViolet, Color.Blue, ShapeType.TextCascadeDown);
            shape.TextPath.Spacing = 0.1;

            // Set the "RotateLetters" property to "true" to rotate each character 90 degrees counterclockwise.
            shape = AppendWordArt(doc, "RotateLetters", "Calibri", 200, 36, Color.GreenYellow, Color.Green, ShapeType.TextWave);
            shape.TextPath.RotateLetters = true;

            // Set the "SameLetterHeights" property to "true" to get the x-height of each character to equal the cap height.
            shape = AppendWordArt(doc, "Same character height for lower and UPPER case", "Calibri", 300, 24, Color.DeepSkyBlue, Color.DodgerBlue, ShapeType.TextSlantUp);
            shape.TextPath.SameLetterHeights = true;

            // By default, the text's size will always scale to fit the containing shape's size, overriding the text size setting.
            shape = AppendWordArt(doc, "FitShape on", "Calibri", 160, 24, Color.LightBlue, Color.Blue, ShapeType.TextPlainText);
            Assert.That(shape.TextPath.FitShape, Is.True);
            shape.TextPath.Size = 24.0;

            // If we set the "FitShape: property to "false", the text will keep the size
            // which the "Size" property specifies regardless of the size of the shape.
            // Use the "TextPathAlignment" property also to align the text to a side of the shape.
            shape = AppendWordArt(doc, "FitShape off", "Calibri", 160, 24, Color.LightBlue, Color.Blue, ShapeType.TextPlainText);
            shape.TextPath.FitShape = false;
            shape.TextPath.Size = 24.0;
            shape.TextPath.TextPathAlignment = TextPathAlignment.Right;

            doc.Save(ArtifactsDir + "Shape.InsertTextPaths.docx");
            TestInsertTextPaths(ArtifactsDir + "Shape.InsertTextPaths.docx"); //ExSkip
        }

        /// <summary>
        /// Insert a new paragraph with a WordArt shape inside it.
        /// </summary>
        private static Shape AppendWordArt(Document doc, string text, string textFontFamily, double shapeWidth, double shapeHeight, Color wordArtFill, Color line, ShapeType wordArtShapeType)
        {
            // Create an inline Shape, which will serve as a container for our WordArt.
            // The shape can only be a valid WordArt shape if we assign a WordArt-designated ShapeType to it.
            // These types will have "WordArt object" in the description,
            // and their enumerator constant names will all start with "Text".
            Shape shape = new Shape(doc, wordArtShapeType)
            {
                WrapType = WrapType.Inline,
                Width = shapeWidth,
                Height = shapeHeight,
                FillColor = wordArtFill,
                StrokeColor = line
            };

            shape.TextPath.Text = text;
            shape.TextPath.FontFamily = textFontFamily;

            Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
            para.AppendChild(shape);
            return shape;
        }
        //ExEnd

        private void TestInsertTextPaths(string filename)
        {
            Document doc = new Document(filename);
            List<Shape> shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToList();

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 480, 24, 0.0d, 0.0d, shapes[0]);
            Assert.That(shapes[0].TextPath.Bold, Is.True);
            Assert.That(shapes[0].TextPath.Italic, Is.True);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 150, 24, 0.0d, 0.0d, shapes[1]);
            Assert.That(shapes[1].TextPath.On, Is.True);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 150, 24, 0.0d, 0.0d, shapes[2]);
            Assert.That(shapes[2].TextPath.On, Is.False);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 90, 24, 0.0d, 0.0d, shapes[3]);
            Assert.That(shapes[3].TextPath.Kerning, Is.True);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 100, 24, 0.0d, 0.0d, shapes[4]);
            Assert.That(shapes[4].TextPath.Kerning, Is.False);

            TestUtil.VerifyShape(ShapeType.TextCascadeDown, string.Empty, 120, 24, 0.0d, 0.0d, shapes[5]);
            Assert.That(shapes[5].TextPath.Spacing, Is.EqualTo(0.1d).Within(0.01d));

            TestUtil.VerifyShape(ShapeType.TextWave, string.Empty, 200, 36, 0.0d, 0.0d, shapes[6]);
            Assert.That(shapes[6].TextPath.RotateLetters, Is.True);

            TestUtil.VerifyShape(ShapeType.TextSlantUp, string.Empty, 300, 24, 0.0d, 0.0d, shapes[7]);
            Assert.That(shapes[7].TextPath.SameLetterHeights, Is.True);

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 160, 24, 0.0d, 0.0d, shapes[8]);
            Assert.That(shapes[8].TextPath.FitShape, Is.True);
            Assert.That(shapes[8].TextPath.Size, Is.EqualTo(24.0d));

            TestUtil.VerifyShape(ShapeType.TextPlainText, string.Empty, 160, 24, 0.0d, 0.0d, shapes[9]);
            Assert.That(shapes[9].TextPath.FitShape, Is.False);
            Assert.That(shapes[9].TextPath.Size, Is.EqualTo(24.0d));
            Assert.That(shapes[9].TextPath.TextPathAlignment, Is.EqualTo(TextPathAlignment.Right));
        }

        [Test]
        public void ShapeRevision()
        {
            //ExStart
            //ExFor:ShapeBase.IsDeleteRevision
            //ExFor:ShapeBase.IsInsertRevision
            //ExSummary:Shows how to work with revision shapes.
            Document doc = new Document();

            Assert.That(doc.TrackRevisions, Is.False);

            // Insert an inline shape without tracking revisions, which will make this shape not a revision of any kind.
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.WrapType = WrapType.Inline;
            shape.Width = 100.0;
            shape.Height = 100.0;
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            // Start tracking revisions and then insert another shape, which will be a revision.
            doc.StartTrackRevisions("John Doe");

            shape = new Shape(doc, ShapeType.Sun);
            shape.WrapType = WrapType.Inline;
            shape.Width = 100.0;
            shape.Height = 100.0;
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.That(shapes.Length, Is.EqualTo(2));

            shapes[0].Remove();

            // Since we removed that shape while we were tracking changes,
            // the shape persists in the document and counts as a delete revision.
            // Accepting this revision will remove the shape permanently, and rejecting it will keep it in the document.
            Assert.That(shapes[0].ShapeType, Is.EqualTo(ShapeType.Cube));
            Assert.That(shapes[0].IsDeleteRevision, Is.True);

            // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
            // Accepting this revision will assimilate this shape into the document as a non-revision,
            // and rejecting the revision will remove this shape permanently.
            Assert.That(shapes[1].ShapeType, Is.EqualTo(ShapeType.Sun));
            Assert.That(shapes[1].IsInsertRevision, Is.True);
            //ExEnd
        }

        [Test]
        public void MoveRevisions()
        {
            //ExStart
            //ExFor:ShapeBase.IsMoveFromRevision
            //ExFor:ShapeBase.IsMoveToRevision
            //ExSummary:Shows how to identify move revision shapes.
            // A move revision is when we move an element in the document body by cut-and-pasting it in Microsoft Word while
            // tracking changes. If we involve an inline shape in such a text movement, that shape will also be a revision.
            // Copying-and-pasting or moving floating shapes do not create move revisions.
            Document doc = new Document(MyDir + "Revision shape.docx");

            // Move revisions consist of pairs of "Move from", and "Move to" revisions. We moved in this document in one shape,
            // but until we accept or reject the move revision, there will be two instances of that shape.
            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.That(shapes.Length, Is.EqualTo(2));

            // This is the "Move to" revision, which is the shape at its arrival destination.
            // If we accept the revision, this "Move to" revision shape will disappear,
            // and the "Move from" revision shape will remain.
            Assert.That(shapes[0].IsMoveFromRevision, Is.False);
            Assert.That(shapes[0].IsMoveToRevision, Is.True);

            // This is the "Move from" revision, which is the shape at its original location.
            // If we accept the revision, this "Move from" revision shape will disappear,
            // and the "Move to" revision shape will remain.
            Assert.That(shapes[1].IsMoveFromRevision, Is.True);
            Assert.That(shapes[1].IsMoveToRevision, Is.False);
            //ExEnd
        }

        [Test]
        public void AdjustWithEffects()
        {
            //ExStart
            //ExFor:ShapeBase.AdjustWithEffects(RectangleF)
            //ExFor:ShapeBase.BoundsWithEffects
            //ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
            Document doc = new Document(MyDir + "Shape shadow effect.docx");

            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.That(shapes.Length, Is.EqualTo(2));

            // The two shapes are identical in terms of dimensions and shape type.
            Assert.That(shapes[1].Width, Is.EqualTo(shapes[0].Width));
            Assert.That(shapes[1].Height, Is.EqualTo(shapes[0].Height));
            Assert.That(shapes[1].ShapeType, Is.EqualTo(shapes[0].ShapeType));

            // The first shape has no effects, and the second one has a shadow and thick outline.
            // These effects make the size of the second shape's silhouette bigger than that of the first.
            // Even though the rectangle's size shows up when we click on these shapes in Microsoft Word,
            // the visible outer bounds of the second shape are affected by the shadow and outline and thus are bigger.
            // We can use the "AdjustWithEffects" method to see the true size of the shape.
            Assert.That(shapes[0].StrokeWeight, Is.EqualTo(0.0));
            Assert.That(shapes[1].StrokeWeight, Is.EqualTo(20.0));
            Assert.That(shapes[0].ShadowEnabled, Is.False);
            Assert.That(shapes[1].ShadowEnabled, Is.True);

            Shape shape = shapes[0];

            // Create a RectangleF object, representing a rectangle,
            // which we could potentially use as the coordinates and bounds for a shape.
            RectangleF rectangleF = new RectangleF(200, 200, 1000, 1000);

            // Run this method to get the size of the rectangle adjusted for all our shape effects.
            RectangleF rectangleFOut = shape.AdjustWithEffects(rectangleF);

            // Since the shape has no border-changing effects, its boundary dimensions are unaffected.
            Assert.That(rectangleFOut.X, Is.EqualTo(200));
            Assert.That(rectangleFOut.Y, Is.EqualTo(200));
            Assert.That(rectangleFOut.Width, Is.EqualTo(1000));
            Assert.That(rectangleFOut.Height, Is.EqualTo(1000));

            // Verify the final extent of the first shape, in points.
            Assert.That(shape.BoundsWithEffects.X, Is.EqualTo(0));
            Assert.That(shape.BoundsWithEffects.Y, Is.EqualTo(0));
            Assert.That(shape.BoundsWithEffects.Width, Is.EqualTo(147));
            Assert.That(shape.BoundsWithEffects.Height, Is.EqualTo(147));

            shape = shapes[1];
            rectangleF = new RectangleF(200, 200, 1000, 1000);
            rectangleFOut = shape.AdjustWithEffects(rectangleF);

            // The shape effects have moved the apparent top left corner of the shape slightly.
            Assert.That(rectangleFOut.X, Is.EqualTo(171.5));
            Assert.That(rectangleFOut.Y, Is.EqualTo(167));

            // The effects have also affected the visible dimensions of the shape.
            Assert.That(rectangleFOut.Width, Is.EqualTo(1045));
            Assert.That(rectangleFOut.Height, Is.EqualTo(1133.5));

            // The effects have also affected the visible bounds of the shape.
            Assert.That(shape.BoundsWithEffects.X, Is.EqualTo(-28.5));
            Assert.That(shape.BoundsWithEffects.Y, Is.EqualTo(-33));
            Assert.That(shape.BoundsWithEffects.Width, Is.EqualTo(192));
            Assert.That(shape.BoundsWithEffects.Height, Is.EqualTo(280.5));
            //ExEnd
        }

        [Test]
        public void RenderAllShapes()
        {
            //ExStart
            //ExFor:ShapeBase.GetShapeRenderer
            //ExFor:NodeRendererBase.Save(Stream, ImageSaveOptions)
            //ExSummary:Shows how to use a shape renderer to export shapes to files in the local file system.
            Document doc = new Document(MyDir + "Various shapes.docx");
            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            Assert.That(shapes.Length, Is.EqualTo(7));

            // There are 7 shapes in the document, including one group shape with 2 child shapes.
            // We will render every shape to an image file in the local file system
            // while ignoring the group shapes since they have no appearance.
            // This will produce 6 image files.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            {
                ShapeRenderer renderer = shape.GetShapeRenderer();
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
                renderer.Save(ArtifactsDir + $"Shape.RenderAllShapes.{shape.Name}.png", options);
            }
            //ExEnd
        }

        [Test]
        public void DocumentHasSmartArtObject()
        {
            //ExStart
            //ExFor:Shape.HasSmartArt
            //ExSummary:Shows how to count the number of shapes in a document with SmartArt objects.
            Document doc = new Document(MyDir + "SmartArt.docx");

            int numberOfSmartArtShapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Count(shape => shape.HasSmartArt);

            Assert.That(numberOfSmartArtShapes, Is.EqualTo(2));
            //ExEnd

        }

        [Test, Category("SkipMono")]
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
            //ExFor:OfficeMathRenderer.#ctor(OfficeMath)
            //ExSummary:Shows how to measure and scale shapes.
            Document doc = new Document(MyDir + "Office math.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            OfficeMathRenderer renderer = new OfficeMathRenderer(officeMath);

            // Verify the size of the image that the OfficeMath object will create when we render it.
            Assert.That(renderer.SizeInPoints.Width, Is.EqualTo(122.0f).Within(0.25f));
            Assert.That(renderer.SizeInPoints.Height, Is.EqualTo(13.0f).Within(0.15f));

            Assert.That(renderer.BoundsInPoints.Width, Is.EqualTo(122.0f).Within(0.25f));
            Assert.That(renderer.BoundsInPoints.Height, Is.EqualTo(13.0f).Within(0.15f));

            // Shapes with transparent parts may contain different values in the "OpaqueBoundsInPoints" properties.
            Assert.That(renderer.OpaqueBoundsInPoints.Width, Is.EqualTo(122.0f).Within(0.25f));
            Assert.That(renderer.OpaqueBoundsInPoints.Height, Is.EqualTo(14.2f).Within(0.1f));

            // Get the shape size in pixels, with linear scaling to a specific DPI.
            Rectangle bounds = renderer.GetBoundsInPixels(1.0f, 96.0f);

            Assert.That(bounds.Width, Is.EqualTo(163));
            Assert.That(bounds.Height, Is.EqualTo(18));

            // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions.
            bounds = renderer.GetBoundsInPixels(1.0f, 96.0f, 150.0f);
            Assert.That(bounds.Width, Is.EqualTo(163));
            Assert.That(bounds.Height, Is.EqualTo(27));

            // The opaque bounds may vary here also.
            bounds = renderer.GetOpaqueBoundsInPixels(1.0f, 96.0f);

            Assert.That(bounds.Width, Is.EqualTo(163));
            Assert.That(bounds.Height, Is.EqualTo(19));

            bounds = renderer.GetOpaqueBoundsInPixels(1.0f, 96.0f, 150.0f);

            Assert.That(bounds.Width, Is.EqualTo(163));
            Assert.That(bounds.Height, Is.EqualTo(29));
            //ExEnd
        }

        [Test]
        public void ShapeTypes()
        {
            //ExStart
            //ExFor:ShapeType
            //ExSummary:Shows how Aspose.Words identify shapes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertShape(ShapeType.Heptagon, RelativeHorizontalPosition.Page, 0,
                RelativeVerticalPosition.Page, 0, 0, 0, WrapType.None);

            builder.InsertShape(ShapeType.Cloud, RelativeHorizontalPosition.RightMargin, 0,
                RelativeVerticalPosition.Page, 0, 0, 0, WrapType.None);

            builder.InsertShape(ShapeType.MathPlus, RelativeHorizontalPosition.RightMargin, 0,
                RelativeVerticalPosition.Page, 0, 0, 0, WrapType.None);

            // To correct identify shape types you need to work with shapes as DML.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                // "Strict" or "Transitional" compliance allows to save shape as DML.
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional
            };

            doc.Save(ArtifactsDir + "Shape.ShapeTypes.docx", saveOptions);
            doc = new Document(ArtifactsDir + "Shape.ShapeTypes.docx");

            Shape[] shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ToArray();

            foreach (Shape shape in shapes)
            {
                Console.WriteLine(shape.ShapeType);
            }
            //ExEnd
        }

        [Test]
        public void IsDecorative()
        {
            //ExStart
            //ExFor:ShapeBase.IsDecorative
            //ExSummary:Shows how to set that the shape is decorative.
            Document doc = new Document(MyDir + "Decorative shapes.docx");

            Shape shape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
            Assert.That(shape.IsDecorative, Is.True);

            // If "AlternativeText" is not empty, the shape cannot be decorative.
            // That's why our value has changed to 'false'.
            shape.AlternativeText = "Alternative text.";
            Assert.That(shape.IsDecorative, Is.False);

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            // Create a new shape as decorative.
            shape = builder.InsertShape(ShapeType.Rectangle, 100, 100);
            shape.IsDecorative = true;

            doc.Save(ArtifactsDir + "Shape.IsDecorative.docx");
            //ExEnd
        }

        [Test]
        public void FillImage()
        {
            //ExStart
            //ExFor:Fill.SetImage(String)
            //ExFor:Fill.SetImage(Byte[])
            //ExFor:Fill.SetImage(Stream)
            //ExSummary:Shows how to set shape fill type as image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // There are several ways of setting image.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 80, 80);
            // 1 -  Using a local system filename:
            shape.Fill.SetImage(ImageDir + "Logo.jpg");
            doc.Save(ArtifactsDir + "Shape.FillImage.FileName.docx");

            // 2 -  Load a file into a byte array:
            shape.Fill.SetImage(File.ReadAllBytes(ImageDir + "Logo.jpg"));
            doc.Save(ArtifactsDir + "Shape.FillImage.ByteArray.docx");

            // 3 -  From a stream:
            using (FileStream stream = new FileStream(ImageDir + "Logo.jpg", FileMode.Open))
                shape.Fill.SetImage(stream);
            doc.Save(ArtifactsDir + "Shape.FillImage.Stream.docx");
            //ExEnd
        }

        [Test]
        public void ShadowFormat()
        {
            //ExStart
            //ExFor:ShadowFormat.Visible
            //ExFor:ShadowFormat.Clear()
            //ExFor:ShadowType
            //ExSummary:Shows how to work with a shadow formatting for the shape.
            Document doc = new Document(MyDir + "Shape stroke pattern border.docx");
            Shape shape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];

            if (shape.ShadowFormat.Visible && shape.ShadowFormat.Type == ShadowType.Shadow2)
                shape.ShadowFormat.Type = ShadowType.Shadow7;

            if (shape.ShadowFormat.Type == ShadowType.ShadowMixed)
                shape.ShadowFormat.Clear();
            //ExEnd
        }

        [Test]
        public void NoTextRotation()
        {
            //ExStart
            //ExFor:TextBox.NoTextRotation
            //ExSummary:Shows how to disable text rotation when the shape is rotate.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.Ellipse, 20, 20);
            shape.TextBox.NoTextRotation = true;

            doc.Save(ArtifactsDir + "Shape.NoTextRotation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Shape.NoTextRotation.docx");
            shape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];

            Assert.That(shape.TextBox.NoTextRotation, Is.EqualTo(true));
        }

        [Test]
        public void RelativeSizeAndPosition()
        {
            //ExStart
            //ExFor:ShapeBase.RelativeHorizontalSize
            //ExFor:ShapeBase.RelativeVerticalSize
            //ExFor:ShapeBase.WidthRelative
            //ExFor:ShapeBase.HeightRelative
            //ExFor:ShapeBase.TopRelative
            //ExFor:ShapeBase.LeftRelative
            //ExFor:RelativeHorizontalSize
            //ExFor:RelativeVerticalSize
            //ExSummary:Shows how to set relative size and position.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Adding a simple shape with absolute size and position.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 40);
            // Set WrapType to WrapType.None since Inline shapes are automatically converted to absolute units.
            shape.WrapType = WrapType.None;

            // Checking and setting the relative horizontal size.
            if (shape.RelativeHorizontalSize == RelativeHorizontalSize.Default)
            {
                // Setting the horizontal size binding to Margin.
                shape.RelativeHorizontalSize = RelativeHorizontalSize.Margin;
                // Setting the width to 50% of Margin width.
                shape.WidthRelative = 50;
            }

            // Checking and setting the relative vertical size.
            if (shape.RelativeVerticalSize == RelativeVerticalSize.Default)
            {
                // Setting the vertical size binding to Margin.
                shape.RelativeVerticalSize = RelativeVerticalSize.Margin;
                // Setting the heigh to 30% of Margin height.
                shape.HeightRelative = 30;
            }

            // Checking and setting the relative vertical position.
            if (shape.RelativeVerticalPosition == RelativeVerticalPosition.Paragraph)
            {
                // etting the position binding to TopMargin.
                shape.RelativeVerticalPosition = RelativeVerticalPosition.TopMargin;
                // Setting relative Top to 30% of TopMargin position.
                shape.TopRelative = 30;
            }

            // Checking and setting the relative horizontal position.
            if (shape.RelativeHorizontalPosition == RelativeHorizontalPosition.Default)
            {
                // Setting the position binding to RightMargin.
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.RightMargin;
                // The position relative value can be negative.
                shape.LeftRelative = -260;
            }

            doc.Save(ArtifactsDir + "Shape.RelativeSizeAndPosition.docx");
            //ExEnd
        }

        [Test]
        public void FillBaseColor()
        {
            //ExStart:FillBaseColor
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:Fill.BaseForeColor
            //ExFor:Stroke.BaseForeColor
            //ExSummary:Shows how to get foreground color without modifiers.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 40);
            shape.Fill.ForeColor = Color.Red;
            shape.Fill.ForeTintAndShade = 0.5;
            shape.Stroke.Fill.ForeColor = Color.Green;
            shape.Stroke.Fill.Transparency = 0.5;

            Assert.That(shape.Fill.ForeColor.ToArgb(), Is.EqualTo(Color.FromArgb(255, 255, 188, 188).ToArgb()));
            Assert.That(shape.Fill.BaseForeColor.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));

            Assert.That(shape.Stroke.ForeColor.ToArgb(), Is.EqualTo(Color.FromArgb(128, 0, 128, 0).ToArgb()));
            Assert.That(shape.Stroke.BaseForeColor.ToArgb(), Is.EqualTo(Color.Green.ToArgb()));

            Assert.That(shape.Stroke.Fill.ForeColor.ToArgb(), Is.EqualTo(Color.Green.ToArgb()));
            Assert.That(shape.Stroke.Fill.BaseForeColor.ToArgb(), Is.EqualTo(Color.Green.ToArgb()));
            //ExEnd:FillBaseColor
        }

        [Test]
        public void FitImageToShape()
        {
            //ExStart:FitImageToShape
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:ImageData.FitImageToShape
            //ExSummary:Shows hot to fit the image data to Shape frame.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an image shape and leave its orientation in its default state.
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 300, 450);
            shape.ImageData.SetImage(ImageDir + "Barcode.png");
            shape.ImageData.FitImageToShape();

            doc.Save(ArtifactsDir + "Shape.FitImageToShape.docx");
            //ExEnd:FitImageToShape
        }

        [Test]
        public void StrokeForeThemeColors()
        {
            //ExStart:StrokeForeThemeColors
            //GistId:eeeec1fbf118e95e7df3f346c91ed726
            //ExFor:Stroke.ForeThemeColor
            //ExFor:Stroke.ForeTintAndShade
            //ExSummary:Shows how to set fore theme color and tint and shade.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.TextBox, 100, 40);
            Stroke stroke = shape.Stroke;
            stroke.ForeThemeColor = ThemeColor.Dark1;
            stroke.ForeTintAndShade = 0.5;

            doc.Save(ArtifactsDir + "Shape.StrokeForeThemeColors.docx");
            //ExEnd:StrokeForeThemeColors

            doc = new Document(ArtifactsDir + "Shape.StrokeForeThemeColors.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.Stroke.ForeThemeColor, Is.EqualTo(ThemeColor.Dark1));
            Assert.That(shape.Stroke.ForeTintAndShade, Is.EqualTo(0.5));
        }

        [Test]
        public void StrokeBackThemeColors()
        {
            //ExStart:StrokeBackThemeColors
            //GistId:eeeec1fbf118e95e7df3f346c91ed726
            //ExFor:Stroke.BackThemeColor
            //ExFor:Stroke.BackTintAndShade
            //ExSummary:Shows how to set back theme color and tint and shade.
            Document doc = new Document(MyDir + "Stroke gradient outline.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Stroke stroke = shape.Stroke;
            stroke.BackThemeColor = ThemeColor.Dark2;
            stroke.BackTintAndShade = 0.2d;

            doc.Save(ArtifactsDir + "Shape.StrokeBackThemeColors.docx");
            //ExEnd:StrokeBackThemeColors

            doc = new Document(ArtifactsDir + "Shape.StrokeBackThemeColors.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.Stroke.BackThemeColor, Is.EqualTo(ThemeColor.Dark2));
            double precision = 1e-6;
            Assert.That(shape.Stroke.BackTintAndShade, Is.EqualTo(0.2d).Within(precision));
        }

        [Test]
        public void TextBoxOleControl()
        {
            //ExStart:TextBoxOleControl
            //GistId:eeeec1fbf118e95e7df3f346c91ed726
            //ExFor:TextBoxControl
            //ExFor:TextBoxControl.Text
            //ExFor:TextBoxControl.Type
            //ExSummary:Shows how to change text of the TextBox OLE control.
            Document doc = new Document(MyDir + "Textbox control.docm");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            TextBoxControl textBoxControl = (TextBoxControl)shape.OleFormat.OleControl;
            Assert.That(textBoxControl.Text, Is.EqualTo("Aspose.Words test"));

            textBoxControl.Text = "Updated text";
            Assert.That(textBoxControl.Text, Is.EqualTo("Updated text"));
            Assert.That(textBoxControl.Type, Is.EqualTo(Forms2OleControlType.Textbox));
            //ExEnd:TextBoxOleControl
        }

        [Test]
        public void Glow()
        {
            //ExStart:Glow
            //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
            //ExFor:ShapeBase.Glow
            //ExFor:GlowFormat
            //ExFor:GlowFormat.Color
            //ExFor:GlowFormat.Radius
            //ExFor:GlowFormat.Transparency
            //ExFor:GlowFormat.Remove()
            //ExSummary:Shows how to interact with glow shape effect.
            Document doc = new Document(MyDir + "Various shapes.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            shape.Glow.Color = Color.Salmon;
            shape.Glow.Radius = 30;
            shape.Glow.Transparency = 0.15;

            doc.Save(ArtifactsDir + "Shape.Glow.docx");

            doc = new Document(ArtifactsDir + "Shape.Glow.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(shape.Glow.Color.ToArgb(), Is.EqualTo(Color.FromArgb(217, 250, 128, 114).ToArgb()));
            Assert.That(shape.Glow.Radius, Is.EqualTo(30));
            Assert.That(shape.Glow.Transparency, Is.EqualTo(0.15d).Within(0.01d));

            shape.Glow.Remove();

            Assert.That(shape.Glow.Color.ToArgb(), Is.EqualTo(Color.Black.ToArgb()));
            Assert.That(shape.Glow.Radius, Is.EqualTo(0));
            Assert.That(shape.Glow.Transparency, Is.EqualTo(0));
            //ExEnd:Glow
        }

        [Test]
        public void Reflection()
        {
            //ExStart:Reflection
            //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
            //ExFor:ShapeBase.Reflection
            //ExFor:ReflectionFormat
            //ExFor:ReflectionFormat.Size
            //ExFor:ReflectionFormat.Blur
            //ExFor:ReflectionFormat.Transparency
            //ExFor:ReflectionFormat.Distance
            //ExFor:ReflectionFormat.Remove()
            //ExSummary:Shows how to interact with reflection shape effect.
            Document doc = new Document(MyDir + "Various shapes.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            shape.Reflection.Transparency = 0.37;
            shape.Reflection.Size = 0.48;
            shape.Reflection.Blur = 17.5;
            shape.Reflection.Distance = 9.2;

            doc.Save(ArtifactsDir + "Shape.Reflection.docx");

            doc = new Document(ArtifactsDir + "Shape.Reflection.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            ReflectionFormat reflectionFormat = shape.Reflection;

            Assert.That(reflectionFormat.Transparency, Is.EqualTo(0.37d).Within(0.01d));
            Assert.That(reflectionFormat.Size, Is.EqualTo(0.48d).Within(0.01d));
            Assert.That(reflectionFormat.Blur, Is.EqualTo(17.5d).Within(0.01d));
            Assert.That(reflectionFormat.Distance, Is.EqualTo(9.2d).Within(0.01d));

            reflectionFormat.Remove();

            Assert.That(reflectionFormat.Transparency, Is.EqualTo(0));
            Assert.That(reflectionFormat.Size, Is.EqualTo(0));
            Assert.That(reflectionFormat.Blur, Is.EqualTo(0));
            Assert.That(reflectionFormat.Distance, Is.EqualTo(0));
            //ExEnd:Reflection
        }

        [Test]
        public void SoftEdge()
        {
            //ExStart:SoftEdge
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:ShapeBase.SoftEdge
            //ExFor:SoftEdgeFormat
            //ExFor:SoftEdgeFormat.Radius
            //ExFor:SoftEdgeFormat.Remove
            //ExSummary:Shows how to work with soft edge formatting.
            DocumentBuilder builder = new DocumentBuilder();
            Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 200);

            // Apply soft edge to the shape.
            shape.SoftEdge.Radius = 30;

            builder.Document.Save(ArtifactsDir + "Shape.SoftEdge.docx");

            // Load document with rectangle shape with soft edge.
            Document doc = new Document(ArtifactsDir + "Shape.SoftEdge.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            SoftEdgeFormat softEdgeFormat = shape.SoftEdge;

            // Check soft edge radius.
            Assert.That(softEdgeFormat.Radius, Is.EqualTo(30));

            // Remove soft edge from the shape.
            softEdgeFormat.Remove();

            // Check radius of the removed soft edge.
            Assert.That(softEdgeFormat.Radius, Is.EqualTo(0));
            //ExEnd:SoftEdge
        }

        [Test]
        public void Adjustments()
        {
            //ExStart:Adjustments
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:Shape.Adjustments
            //ExFor:AdjustmentCollection
            //ExFor:AdjustmentCollection.Count
            //ExFor:AdjustmentCollection.Item(Int32)
            //ExFor:Adjustment
            //ExFor:Adjustment.Name
            //ExFor:Adjustment.Value
            //ExSummary:Shows how to work with adjustment raw values.
            Document doc = new Document(MyDir + "Rounded rectangle shape.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            AdjustmentCollection adjustments = shape.Adjustments;
            Assert.That(adjustments.Count, Is.EqualTo(1));

            Adjustment adjustment = adjustments[0];
            Assert.That(adjustment.Name, Is.EqualTo("adj"));
            Assert.That(adjustment.Value, Is.EqualTo(16667));

            adjustment.Value = 30000;

            doc.Save(ArtifactsDir + "Shape.Adjustments.docx");
            //ExEnd:Adjustments

            doc = new Document(ArtifactsDir + "Shape.Adjustments.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            adjustments = shape.Adjustments;
            Assert.That(adjustments.Count, Is.EqualTo(1));

            adjustment = adjustments[0];
            Assert.That(adjustment.Name, Is.EqualTo("adj"));
            Assert.That(adjustment.Value, Is.EqualTo(30000));
        }

        [Test]
        public void ShadowFormatColor()
        {
            //ExStart:ShadowFormatColor
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ShapeBase.ShadowFormat
            //ExFor:ShadowFormat
            //ExFor:ShadowFormat.Color
            //ExFor:ShadowFormat.Type
            //ExSummary:Shows how to get shadow color.
            Document doc = new Document(MyDir + "Shadow color.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            ShadowFormat shadowFormat = shape.ShadowFormat;

            Assert.That(shadowFormat.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            Assert.That(shadowFormat.Type, Is.EqualTo(ShadowType.ShadowMixed));
            //ExEnd:ShadowFormatColor
        }

        [Test]
        public void SetActiveXProperties()
        {
            //ExStart:SetActiveXProperties
            //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
            //ExFor:Forms2OleControl.ForeColor
            //ExFor:Forms2OleControl.BackColor
            //ExFor:Forms2OleControl.Height
            //ExFor:Forms2OleControl.Width
            //ExSummary:Shows how to set properties for ActiveX control.
            Document doc = new Document(MyDir + "ActiveX controls.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Forms2OleControl oleControl = (Forms2OleControl)shape.OleFormat.OleControl;
            oleControl.ForeColor = Color.FromArgb(0x17, 0xE1, 0x35);
            oleControl.BackColor = Color.FromArgb(0x33, 0x97, 0xF4);
            oleControl.Height = 100.54;
            oleControl.Width = 201.06;
            //ExEnd:SetActiveXProperties

            Assert.That(oleControl.ForeColor.ToArgb(), Is.EqualTo(Color.FromArgb(0x17, 0xE1, 0x35).ToArgb()));
            Assert.That(oleControl.BackColor.ToArgb(), Is.EqualTo(Color.FromArgb(0x33, 0x97, 0xF4).ToArgb()));
            Assert.That(oleControl.Height, Is.EqualTo(100.54));
            Assert.That(oleControl.Width, Is.EqualTo(201.06));
        }

        [Test]
        public void SelectRadioControl()
        {
            //ExStart:SelectRadioControl
            //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
            //ExFor:OptionButtonControl
            //ExFor:OptionButtonControl.Selected
            //ExFor:OptionButtonControl.Type
            //ExSummary:Shows how to select radio button.
            Document doc = new Document(MyDir + "Radio buttons.docx");

            Shape shape1 = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            OptionButtonControl optionButton1 = (OptionButtonControl)shape1.OleFormat.OleControl;
            // Deselect selected first item.
            optionButton1.Selected = false;

            Shape shape2 = (Shape)doc.GetChild(NodeType.Shape, 1, true);
            OptionButtonControl optionButton2 = (OptionButtonControl)shape2.OleFormat.OleControl;
            // Select second option button.
            optionButton2.Selected = true;

            Assert.That(optionButton1.Type, Is.EqualTo(Forms2OleControlType.OptionButton));
            Assert.That(optionButton2.Type, Is.EqualTo(Forms2OleControlType.OptionButton));

            doc.Save(ArtifactsDir + "Shape.SelectRadioControl.docx");
            //ExEnd:SelectRadioControl
        }

        [Test]
        public void CheckedCheckBox()
        {
            //ExStart:CheckedCheckBox
            //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
            //ExFor:CheckBoxControl
            //ExFor:CheckBoxControl.Checked
            //ExFor:CheckBoxControl.Type
            //ExFor:Forms2OleControlType
            //ExSummary:Shows how to change state of the CheckBox control.
            Document doc = new Document(MyDir + "ActiveX controls.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            CheckBoxControl checkBoxControl = (CheckBoxControl)shape.OleFormat.OleControl;
            checkBoxControl.Checked = true;
            
            Assert.That(checkBoxControl.Checked, Is.EqualTo(true));
            Assert.That(checkBoxControl.Type, Is.EqualTo(Forms2OleControlType.CheckBox));
            //ExEnd:CheckedCheckBox
        }

        [Test]
        public void InsertGroupShape()
        {
            //ExStart:InsertGroupShape
            //GistId:e06aa7a168b57907a5598e823a22bf0a
            //ExFor:DocumentBuilder.InsertGroupShape(double, double, double, double, ShapeBase[])
            //ExFor:DocumentBuilder.InsertGroupShape(ShapeBase[])
            //ExSummary:Shows how to insert DML group shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
            shape1.Left = 20;
            shape1.Top = 20;
            shape1.Stroke.Color = Color.Red;

            Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
            shape2.Left = 40;
            shape2.Top = 50;
            shape2.Stroke.Color = Color.Green;

            // Dimensions for the new GroupShape node.
            double left = 10;
            double top = 10;
            double width = 200;
            double height = 300;
            // Insert GroupShape node for the specified size which is inserted into the specified position.
            GroupShape groupShape1 = builder.InsertGroupShape(left, top, width, height, new Shape[] { shape1, shape2 });

            // Insert GroupShape node which position and dimension will be calculated automatically.
            Shape shape3 = (Shape)shape1.Clone(true);
            GroupShape groupShape2 = builder.InsertGroupShape(shape3);

            doc.Save(ArtifactsDir + "Shape.InsertGroupShape.docx");
            //ExEnd:InsertGroupShape
        }

        [Test]
        public void CombineGroupShape()
        {
            //ExStart:CombineGroupShape
            //GistId:bb594993b5fe48692541e16f4d354ac2
            //ExFor:DocumentBuilder.InsertGroupShape(ShapeBase[])
            //ExSummary:Shows how to combine group shape with the shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
            shape1.Left = 20;
            shape1.Top = 20;
            shape1.Stroke.Color = Color.Red;

            Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
            shape2.Left = 40;
            shape2.Top = 50;
            shape2.Stroke.Color = Color.Green;

            // Combine shapes into a GroupShape node which is inserted into the specified position.
            GroupShape groupShape1 = builder.InsertGroupShape(shape1, shape2);

            // Combine Shape and GroupShape nodes.
            Shape shape3 = (Shape)shape1.Clone(true);
            GroupShape groupShape2 = builder.InsertGroupShape(groupShape1, shape3);

            doc.Save(ArtifactsDir + "Shape.CombineGroupShape.docx");
            //ExEnd:CombineGroupShape

            doc = new Document(ArtifactsDir + "Shape.CombineGroupShape.docx");

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                Assert.That(shape.Width, Is.Not.EqualTo(0));
                Assert.That(shape.Height, Is.Not.EqualTo(0));
            }
        }

        [Test]
        public void InsertCommandButton()
        {
            //ExStart:InsertCommandButton
            //GistId:bb594993b5fe48692541e16f4d354ac2
            //ExFor:CommandButtonControl
            //ExFor:CommandButtonControl.#ctor
            //ExFor:CommandButtonControl.Type
            //ExFor:DocumentBuilder.InsertForms2OleControl(Forms2OleControl)
            //ExSummary:Shows how to insert ActiveX control.
            DocumentBuilder builder = new DocumentBuilder();

            CommandButtonControl button1 = new CommandButtonControl();
            Shape shape = builder.InsertForms2OleControl(button1);
            Assert.That(button1.Type, Is.EqualTo(Forms2OleControlType.CommandButton));
            //ExEnd:InsertCommandButton
        }

        [Test]
        public void Hidden()
        {
            //ExStart:Hidden
            //GistId:bb594993b5fe48692541e16f4d354ac2
            //ExFor:ShapeBase.Hidden
            //ExSummary:Shows how to hide the shape.
            Document doc = new Document(MyDir + "Shadow color.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            if (!shape.Hidden)
                shape.Hidden = true;

            doc.Save(ArtifactsDir + "Shape.Hidden.docx");
            //ExEnd:Hidden
        }

        [Test]
        public void CommandButtonCaption()
        {
            //ExStart:CommandButtonCaption
            //GistId:366eb64fd56dec3c2eaa40410e594182
            //ExFor:Forms2OleControl.Caption
            //ExSummary:Shows how to set caption for ActiveX control.
            DocumentBuilder builder = new DocumentBuilder();

            CommandButtonControl button1 = new CommandButtonControl() { Caption = "Button caption" };
            Shape shape = builder.InsertForms2OleControl(button1);
            Assert.That(button1.Caption, Is.EqualTo("Button caption"));
            //ExEnd:CommandButtonCaption
        }
    }
}
