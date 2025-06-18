// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
using NUnit.Framework;
using Shape = Aspose.Words.Drawing.Shape;
#if NET461_OR_GREATER || JAVA
using System.Drawing.Imaging;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDrawing : ApiExampleBase
    {
#if NET461_OR_GREATER || JAVA
        [Test]
        public void VariousShapes()
        {
            //ExStart
            //ExFor:ArrowLength
            //ExFor:ArrowType
            //ExFor:ArrowWidth
            //ExFor:DashStyle
            //ExFor:EndCap
            //ExFor:Fill.ForeColor
            //ExFor:Fill.ImageBytes
            //ExFor:Fill.Visible
            //ExFor:JoinStyle
            //ExFor:Shape.Stroke
            //ExFor:Stroke.Color
            //ExFor:Stroke.StartArrowLength
            //ExFor:Stroke.StartArrowType
            //ExFor:Stroke.StartArrowWidth
            //ExFor:Stroke.EndArrowLength
            //ExFor:Stroke.EndArrowWidth
            //ExFor:Stroke.DashStyle
            //ExFor:Stroke.EndArrowType
            //ExFor:Stroke.EndCap
            //ExFor:Stroke.Opacity
            //ExSummary:Shows to create a variety of shapes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are four examples of shapes that we can insert into our documents.
            // 1 -  Dotted, horizontal, half-transparent red line
            // with an arrow on the left end and a diamond on the right end:
            Shape arrow = new Shape(doc, ShapeType.Line);
            arrow.Width = 200;
            arrow.Stroke.Color = Color.Red;
            arrow.Stroke.StartArrowType = ArrowType.Arrow;
            arrow.Stroke.StartArrowLength = ArrowLength.Long;
            arrow.Stroke.StartArrowWidth = ArrowWidth.Wide;
            arrow.Stroke.EndArrowType = ArrowType.Diamond;
            arrow.Stroke.EndArrowLength = ArrowLength.Long;
            arrow.Stroke.EndArrowWidth = ArrowWidth.Wide;
            arrow.Stroke.DashStyle = DashStyle.Dash;
            arrow.Stroke.Opacity = 0.5;

            Assert.That(arrow.Stroke.JoinStyle, Is.EqualTo(JoinStyle.Miter));

            builder.InsertNode(arrow);

            // 2 -  Thick black diagonal line with rounded ends:
            Shape line = new Shape(doc, ShapeType.Line);
            line.Top = 40;
            line.Width = 200;
            line.Height = 20;
            line.StrokeWeight = 5.0;
            line.Stroke.EndCap = EndCap.Round;

            builder.InsertNode(line);

            // 3 -  Arrow with a green fill:
            Shape filledInArrow = new Shape(doc, ShapeType.Arrow);
            filledInArrow.Width = 200;
            filledInArrow.Height = 40;
            filledInArrow.Top = 100;
            filledInArrow.Fill.ForeColor = Color.Green;
            filledInArrow.Fill.Visible = true;

            builder.InsertNode(filledInArrow);

            // 4 -  Arrow with a flipped orientation filled in with the Aspose logo:
            Shape filledInArrowImg = new Shape(doc, ShapeType.Arrow);
            filledInArrowImg.Width = 200;
            filledInArrowImg.Height = 40;
            filledInArrowImg.Top = 160;
            filledInArrowImg.FlipOrientation = FlipOrientation.Both;

            byte[] imageBytes = File.ReadAllBytes(ImageDir + "Logo.jpg");

            using (MemoryStream stream = new MemoryStream(imageBytes))
            {
                Image image = Image.FromStream(stream);
                // When we flip the orientation of our arrow, we also flip the image that the arrow contains.
                // Flip the image the other way to cancel this out before getting the shape to display it.
                image.RotateFlip(RotateFlipType.RotateNoneFlipXY);

                filledInArrowImg.ImageData.SetImage(image);
                filledInArrowImg.Stroke.JoinStyle = JoinStyle.Round;

                builder.InsertNode(filledInArrowImg);
            }

            doc.Save(ArtifactsDir + "Drawing.VariousShapes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Drawing.VariousShapes.docx");

            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(4));

            arrow = (Shape) doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(arrow.ShapeType, Is.EqualTo(ShapeType.Line));
            Assert.That(arrow.Width, Is.EqualTo(200.0d));
            Assert.That(arrow.Stroke.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            Assert.That(arrow.Stroke.StartArrowType, Is.EqualTo(ArrowType.Arrow));
            Assert.That(arrow.Stroke.StartArrowLength, Is.EqualTo(ArrowLength.Long));
            Assert.That(arrow.Stroke.StartArrowWidth, Is.EqualTo(ArrowWidth.Wide));
            Assert.That(arrow.Stroke.EndArrowType, Is.EqualTo(ArrowType.Diamond));
            Assert.That(arrow.Stroke.EndArrowLength, Is.EqualTo(ArrowLength.Long));
            Assert.That(arrow.Stroke.EndArrowWidth, Is.EqualTo(ArrowWidth.Wide));
            Assert.That(arrow.Stroke.DashStyle, Is.EqualTo(DashStyle.Dash));
            Assert.That(arrow.Stroke.Opacity, Is.EqualTo(0.5d));

            line = (Shape) doc.GetChild(NodeType.Shape, 1, true);

            Assert.That(line.ShapeType, Is.EqualTo(ShapeType.Line));
            Assert.That(line.Top, Is.EqualTo(40.0d));
            Assert.That(line.Width, Is.EqualTo(200.0d));
            Assert.That(line.Height, Is.EqualTo(20.0d));
            Assert.That(line.StrokeWeight, Is.EqualTo(5.0d));
            Assert.That(line.Stroke.EndCap, Is.EqualTo(EndCap.Round));

            filledInArrow = (Shape) doc.GetChild(NodeType.Shape, 2, true);

            Assert.That(filledInArrow.ShapeType, Is.EqualTo(ShapeType.Arrow));
            Assert.That(filledInArrow.Width, Is.EqualTo(200.0d));
            Assert.That(filledInArrow.Height, Is.EqualTo(40.0d));
            Assert.That(filledInArrow.Top, Is.EqualTo(100.0d));
            Assert.That(filledInArrow.Fill.ForeColor.ToArgb(), Is.EqualTo(Color.Green.ToArgb()));
            Assert.That(filledInArrow.Fill.Visible, Is.True);

            filledInArrowImg = (Shape) doc.GetChild(NodeType.Shape, 3, true);

            Assert.That(filledInArrowImg.ShapeType, Is.EqualTo(ShapeType.Arrow));
            Assert.That(filledInArrowImg.Width, Is.EqualTo(200.0d));
            Assert.That(filledInArrowImg.Height, Is.EqualTo(40.0d));
            Assert.That(filledInArrowImg.Top, Is.EqualTo(160.0d));
            Assert.That(filledInArrowImg.FlipOrientation, Is.EqualTo(FlipOrientation.Both));
        }

        [Test]
        public void ImportImage()
        {
            //ExStart
            //ExFor:ImageData.SetImage(Image)
            //ExFor:ImageData.SetImage(Stream)
            //ExSummary:Shows how to display images from the local file system in a document.
            Document doc = new Document();

            // To display an image in a document, we will need to create a shape
            // which will contain an image, and then append it to the document's body.
            Shape imgShape;

            // Below are two ways of getting an image from a file in the local file system.
            // 1 -  Create an image object from an image file:
            using (Image srcImage = Image.FromFile(ImageDir + "Logo.jpg"))
            {
                imgShape = new Shape(doc, ShapeType.Image);
                doc.FirstSection.Body.FirstParagraph.AppendChild(imgShape);
                imgShape.ImageData.SetImage(srcImage);
            }
            
            // 2 -  Open an image file from the local file system using a stream:
            using (Stream stream = new FileStream(ImageDir + "Logo.jpg", FileMode.Open, FileAccess.Read))
            {
                imgShape = new Shape(doc, ShapeType.Image);
                doc.FirstSection.Body.FirstParagraph.AppendChild(imgShape);
                imgShape.ImageData.SetImage(stream);
                imgShape.Left = 150.0f;
            }

            doc.Save(ArtifactsDir + "Drawing.ImportImage.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Drawing.ImportImage.docx");

            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(2));

            imgShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imgShape);
            Assert.That(imgShape.Left, Is.EqualTo(0.0d));
            Assert.That(imgShape.Top, Is.EqualTo(0.0d));
            Assert.That(imgShape.Height, Is.EqualTo(300.0d));
            Assert.That(imgShape.Width, Is.EqualTo(300.0d));
            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imgShape);

            imgShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imgShape);
            Assert.That(imgShape.Left, Is.EqualTo(150.0d));
            Assert.That(imgShape.Top, Is.EqualTo(0.0d));
            Assert.That(imgShape.Height, Is.EqualTo(300.0d));
            Assert.That(imgShape.Width, Is.EqualTo(300.0d));
        }
#endif

        [Test]
        public void TypeOfImage()
        {
            //ExStart
            //ExFor:ImageType
            //ExSummary:Shows how to add an image to a shape and check its type.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape imgShape = builder.InsertImage(ImageDir + "Logo.jpg");
            Assert.That(imgShape.ImageData.ImageType, Is.EqualTo(ImageType.Jpeg));
            //ExEnd
        }

        [Test]
        public void FillSolid()
        {
            //ExStart
            //ExFor:Fill.Color()
            //ExFor:FillType
            //ExFor:Fill.FillType
            //ExFor:Fill.Solid
            //ExFor:Fill.Transparency
            //ExFor:Font.Fill
            //ExSummary:Shows how to convert any of the fills back to solid fill.
            Document doc = new Document(MyDir + "Two color gradient.docx");

            // Get Fill object for Font of the first Run.
            Fill fill = doc.FirstSection.Body.Paragraphs[0].Runs[0].Font.Fill;

            // Check Fill properties of the Font.
            Console.WriteLine("The type of the fill is: {0}", fill.FillType);
            Console.WriteLine("The foreground color of the fill is: {0}", fill.ForeColor);
            Console.WriteLine("The fill is transparent at {0}%", fill.Transparency * 100);

            // Change type of the fill to Solid with uniform green color.
            fill.Solid();
            Console.WriteLine("\nThe fill is changed:");
            Console.WriteLine("The type of the fill is: {0}", fill.FillType);
            Console.WriteLine("The foreground color of the fill is: {0}", fill.ForeColor);
            Console.WriteLine("The fill transparency is {0}%", fill.Transparency * 100);

            doc.Save(ArtifactsDir + "Drawing.FillSolid.docx");
            //ExEnd
        }

        [Test]
        public void SaveAllImages()
        {
            //ExStart
            //ExFor:ImageData.HasImage
            //ExFor:ImageData.ToImage
            //ExFor:ImageData.Save(Stream)
            //ExSummary:Shows how to save all images from a document to the file system.
            Document imgSourceDoc = new Document(MyDir + "Images.docx");

            // Shapes with the "HasImage" flag set store and display all the document's images.
            Shape[] shapesWithImages = imgSourceDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                .Where(s => s.HasImage).ToArray();

            // Go through each shape and save its image.
            for (int shapeIndex = 0; shapeIndex < shapesWithImages.Length; ++shapeIndex)
            {
                ImageData imageData = shapesWithImages[shapeIndex].ImageData;
                using (FileStream fileStream = File.Create(ArtifactsDir + $"Drawing.SaveAllImages.{shapeIndex + 1}.{imageData.ImageType}"))
                    imageData.Save(fileStream);
            }
            //ExEnd

            string[] imageFileNames = Directory.GetFiles(ArtifactsDir).Where(s => s.StartsWith(ArtifactsDir + "Drawing.SaveAllImages.")).OrderBy(s => s).ToArray();
            List<FileInfo> fileInfos = imageFileNames.Select(s => new FileInfo(s)).ToList();
            
            TestUtil.VerifyImage(2467, 1500, fileInfos[0].FullName);
            Assert.That(fileInfos[0].Extension, Is.EqualTo(".Jpeg"));
            TestUtil.VerifyImage(400, 400, fileInfos[1].FullName);
            Assert.That(fileInfos[1].Extension, Is.EqualTo(".Png"));
#if NET461_OR_GREATER || JAVA
            TestUtil.VerifyImage(382, 138, fileInfos[2].FullName);
            Assert.That(fileInfos[2].Extension, Is.EqualTo(".Emf"));
            TestUtil.VerifyImage(1600, 1600, fileInfos[3].FullName);
            Assert.That(fileInfos[3].Extension, Is.EqualTo(".Wmf"));
            TestUtil.VerifyImage(534, 534, fileInfos[4].FullName);
            Assert.That(fileInfos[4].Extension, Is.EqualTo(".Emf"));
#endif
            TestUtil.VerifyImage(1260, 660, fileInfos[5].FullName);
            Assert.That(fileInfos[5].Extension, Is.EqualTo(".Jpeg"));
            TestUtil.VerifyImage(1125, 1500, fileInfos[6].FullName);
            Assert.That(fileInfos[6].Extension, Is.EqualTo(".Jpeg"));
            TestUtil.VerifyImage(1027, 1500, fileInfos[7].FullName);
            Assert.That(fileInfos[7].Extension, Is.EqualTo(".Jpeg"));
            TestUtil.VerifyImage(1200, 1500, fileInfos[8].FullName);
            Assert.That(fileInfos[8].Extension, Is.EqualTo(".Jpeg"));
        }

        [Test]
        public void StrokePattern()
        {
            //ExStart
            //ExFor:Stroke.Color2
            //ExFor:Stroke.ImageBytes
            //ExSummary:Shows how to process shape stroke features.
            Document doc = new Document(MyDir + "Shape stroke pattern border.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Stroke stroke = shape.Stroke;

            // Strokes can have two colors, which are used to create a pattern defined by two-tone image data.
            // Strokes with a single color do not use the Color2 property.
            Assert.That(stroke.Color, Is.EqualTo(Color.FromArgb(255, 128, 0, 0)));
            Assert.That(stroke.Color2, Is.EqualTo(Color.FromArgb(255, 255, 255, 0)));

            Assert.That(stroke.ImageBytes, Is.Not.Null);
            File.WriteAllBytes(ArtifactsDir + "Drawing.StrokePattern.png", stroke.ImageBytes);
            //ExEnd

            TestUtil.VerifyImage(8, 8, ArtifactsDir + "Drawing.StrokePattern.png");
        }

        //ExStart
        //ExFor:DocumentVisitor.VisitShapeEnd(Shape)
        //ExFor:DocumentVisitor.VisitShapeStart(Shape)
        //ExFor:DocumentVisitor.VisitGroupShapeEnd(GroupShape)
        //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
        //ExFor:GroupShape
        //ExFor:GroupShape.#ctor(DocumentBase)
        //ExFor:GroupShape.Accept(DocumentVisitor)
        //ExFor:GroupShape.AcceptStart(DocumentVisitor)
        //ExFor:GroupShape.AcceptEnd(DocumentVisitor)
        //ExFor:ShapeBase.IsGroup
        //ExFor:ShapeBase.ShapeType
        //ExSummary:Shows how to create a group of shapes, and print its contents using a document visitor.
        [Test] //ExSkip
        public void GroupOfShapes()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // If you need to create "NonPrimitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
            // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
            // please use DocumentBuilder.InsertShape methods.
            Shape balloon = new Shape(doc, ShapeType.Balloon)
            {
                Width = 200,
                Height = 200,
                Stroke = { Color = Color.Red }
            };

            Shape cube = new Shape(doc, ShapeType.Cube)
            {
                Width = 100,
                Height = 100,
                Stroke = { Color = Color.Blue }
            };

            GroupShape group = new GroupShape(doc);
            group.AppendChild(balloon);
            group.AppendChild(cube);

            Assert.That(group.IsGroup, Is.True);

            builder.InsertNode(group);

            ShapeGroupPrinter printer = new ShapeGroupPrinter();
            group.Accept(printer);

            Console.WriteLine(printer.GetText());
            TestGroupShapes(doc); //ExSkip
        }

        /// <summary>
        /// Prints the contents of a visited shape group to the console.
        /// </summary>
        public class ShapeGroupPrinter : DocumentVisitor
        {
            public ShapeGroupPrinter()
            {
                mBuilder = new StringBuilder();
            }

            public string GetText()
            {
                return mBuilder.ToString();
            }

            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                mBuilder.AppendLine("Shape group started:");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
            {
                mBuilder.AppendLine("End of shape group");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitShapeStart(Shape shape)
            {
                mBuilder.AppendLine("\tShape - " + shape.ShapeType + ":");
                mBuilder.AppendLine("\t\tWidth: " + shape.Width);
                mBuilder.AppendLine("\t\tHeight: " + shape.Height);
                mBuilder.AppendLine("\t\tStroke color: " + shape.Stroke.Color);
                mBuilder.AppendLine("\t\tFill color: " + shape.Fill.ForeColor);
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitShapeEnd(Shape shape)
            {
                mBuilder.AppendLine("\tEnd of shape");
                return VisitorAction.Continue;
            }

            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private static void TestGroupShapes(Document doc)
        {
            doc = DocumentHelper.SaveOpen(doc);
            GroupShape shapes = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);

            Assert.That(shapes.GetChildNodes(NodeType.Any, false).Count, Is.EqualTo(2));

            Shape shape = (Shape)shapes.GetChildNodes(NodeType.Any, false)[0];

            Assert.That(shape.ShapeType, Is.EqualTo(ShapeType.Balloon));
            Assert.That(shape.Width, Is.EqualTo(200.0d));
            Assert.That(shape.Height, Is.EqualTo(200.0d));
            Assert.That(shape.StrokeColor.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));

            shape = (Shape)shapes.GetChildNodes(NodeType.Any, false)[1];

            Assert.That(shape.ShapeType, Is.EqualTo(ShapeType.Cube));
            Assert.That(shape.Width, Is.EqualTo(100.0d));
            Assert.That(shape.Height, Is.EqualTo(100.0d));
            Assert.That(shape.StrokeColor.ToArgb(), Is.EqualTo(Color.Blue.ToArgb()));
        }

        [Test]
        public void TextBox()
        {
            //ExStart
            //ExFor:LayoutFlow
            //ExSummary:Shows how to add text to a text box, and change its orientation
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textbox = new Shape(doc, ShapeType.TextBox)
            {
                Width = 100,
                Height = 100
            };
            textbox.TextBox.LayoutFlow = LayoutFlow.BottomToTop;

            textbox.AppendChild(new Paragraph(doc));
            builder.InsertNode(textbox);

            builder.MoveTo(textbox.FirstParagraph);
            builder.Write("This text is flipped 90 degrees to the left.");

            doc.Save(ArtifactsDir + "Drawing.TextBox.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Drawing.TextBox.docx");
            textbox = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(textbox.ShapeType, Is.EqualTo(ShapeType.TextBox));
            Assert.That(textbox.Width, Is.EqualTo(100.0d));
            Assert.That(textbox.Height, Is.EqualTo(100.0d));
            Assert.That(textbox.TextBox.LayoutFlow, Is.EqualTo(LayoutFlow.BottomToTop));
            Assert.That(textbox.GetText().Trim(), Is.EqualTo("This text is flipped 90 degrees to the left."));
        }

        [Test]
        public void GetDataFromImage()
        {
            //ExStart
            //ExFor:ImageData.ImageBytes
            //ExFor:ImageData.ToByteArray
            //ExFor:ImageData.ToStream
            //ExSummary:Shows how to create an image file from a shape's raw image data.
            Document imgSourceDoc = new Document(MyDir + "Images.docx");
            Assert.That(imgSourceDoc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(10)); //ExSkip

            Shape imgShape = (Shape) imgSourceDoc.GetChild(NodeType.Shape, 0, true);

            Assert.That(imgShape.HasImage, Is.True);

            // ToByteArray() returns the array stored in the ImageBytes property.
            Assert.That(imgShape.ImageData.ToByteArray(), Is.EqualTo(imgShape.ImageData.ImageBytes));

            // Save the shape's image data to an image file in the local file system.
            using (Stream imgStream = imgShape.ImageData.ToStream())
            {
                using (FileStream outStream = new FileStream(ArtifactsDir + "Drawing.GetDataFromImage.png",
                    FileMode.Create, FileAccess.ReadWrite))
                {
                    imgStream.CopyTo(outStream);
                }
            }
            //ExEnd

            TestUtil.VerifyImage(2467, 1500, ArtifactsDir + "Drawing.GetDataFromImage.png");
        }

        [Test]
        public void ImageData()
        {
            //ExStart
            //ExFor:ImageData.BiLevel
            //ExFor:ImageData.Borders
            //ExFor:ImageData.Brightness
            //ExFor:ImageData.ChromaKey
            //ExFor:ImageData.Contrast
            //ExFor:ImageData.CropBottom
            //ExFor:ImageData.CropLeft
            //ExFor:ImageData.CropRight
            //ExFor:ImageData.CropTop
            //ExFor:ImageData.GrayScale
            //ExFor:ImageData.IsLink
            //ExFor:ImageData.IsLinkOnly
            //ExFor:ImageData.Title
            //ExSummary:Shows how to edit a shape's image data.
            Document imgSourceDoc = new Document(MyDir + "Images.docx");
            Shape sourceShape = (Shape)imgSourceDoc.GetChildNodes(NodeType.Shape, true)[0];

            Document dstDoc = new Document();

            // Import a shape from the source document and append it to the first paragraph.
            Shape importedShape = (Shape)dstDoc.ImportNode(sourceShape, true);
            dstDoc.FirstSection.Body.FirstParagraph.AppendChild(importedShape);

            // The imported shape contains an image. We can access the image's properties and raw data via the ImageData object.
            ImageData imageData = importedShape.ImageData;
            imageData.Title = "Imported Image";

            Assert.That(imageData.HasImage, Is.True);

            // If an image has no borders, its ImageData object will define the border color as empty.
            Assert.That(imageData.Borders.Count, Is.EqualTo(4));
            Assert.That(imageData.Borders[0].Color, Is.EqualTo(Color.Empty));

            // This image does not link to another shape or image file in the local file system.
            Assert.That(imageData.IsLink, Is.False);
            Assert.That(imageData.IsLinkOnly, Is.False);

            // The "Brightness" and "Contrast" properties define image brightness and contrast
            // on a 0-1 scale, with the default value at 0.5.
            imageData.Brightness = 0.8;
            imageData.Contrast = 1.0;

            // The above brightness and contrast values have created an image with a lot of white.
            // We can select a color with the ChromaKey property to replace with transparency, such as white.
            imageData.ChromaKey = Color.White;

            // Import the source shape again and set the image to monochrome.
            importedShape = (Shape)dstDoc.ImportNode(sourceShape, true);
            dstDoc.FirstSection.Body.FirstParagraph.AppendChild(importedShape);

            importedShape.ImageData.GrayScale = true;

            // Import the source shape again to create a third image and set it to BiLevel.
            // BiLevel sets every pixel to either black or white, whichever is closer to the original color.
            importedShape = (Shape)dstDoc.ImportNode(sourceShape, true);
            dstDoc.FirstSection.Body.FirstParagraph.AppendChild(importedShape);

            importedShape.ImageData.BiLevel = true;

            // Cropping is determined on a 0-1 scale. Cropping a side by 0.3
            // will crop 30% of the image out at the cropped side.
            importedShape.ImageData.CropBottom = 0.3;
            importedShape.ImageData.CropLeft = 0.3;
            importedShape.ImageData.CropTop = 0.3;
            importedShape.ImageData.CropRight = 0.3;

            dstDoc.Save(ArtifactsDir + "Drawing.ImageData.docx");
            //ExEnd

            imgSourceDoc = new Document(ArtifactsDir + "Drawing.ImageData.docx");
            sourceShape = (Shape)imgSourceDoc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(2467, 1500, ImageType.Jpeg, sourceShape);
            Assert.That(sourceShape.ImageData.Title, Is.EqualTo("Imported Image"));
            Assert.That(sourceShape.ImageData.Brightness, Is.EqualTo(0.8d).Within(0.1d));
            Assert.That(sourceShape.ImageData.Contrast, Is.EqualTo(1.0d).Within(0.1d));
            Assert.That(sourceShape.ImageData.ChromaKey.ToArgb(), Is.EqualTo(Color.White.ToArgb()));

            sourceShape = (Shape)imgSourceDoc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(2467, 1500, ImageType.Jpeg, sourceShape);
            Assert.That(sourceShape.ImageData.GrayScale, Is.True);

            sourceShape = (Shape)imgSourceDoc.GetChild(NodeType.Shape, 2, true);

            TestUtil.VerifyImageInShape(2467, 1500, ImageType.Jpeg, sourceShape);
            Assert.That(sourceShape.ImageData.BiLevel, Is.True);
            Assert.That(sourceShape.ImageData.CropBottom, Is.EqualTo(0.3d).Within(0.1d));
            Assert.That(sourceShape.ImageData.CropLeft, Is.EqualTo(0.3d).Within(0.1d));
            Assert.That(sourceShape.ImageData.CropTop, Is.EqualTo(0.3d).Within(0.1d));
            Assert.That(sourceShape.ImageData.CropRight, Is.EqualTo(0.3d).Within(0.1d));
        }

        [Test]
        public void ImageSize()
        {
            //ExStart
            //ExFor:ImageSize.HeightPixels
            //ExFor:ImageSize.HorizontalResolution
            //ExFor:ImageSize.VerticalResolution
            //ExFor:ImageSize.WidthPixels
            //ExSummary:Shows how to read the properties of an image in a shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a shape into the document which contains an image taken from our local file system.
            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");

            // If the shape contains an image, its ImageData property will be valid,
            // and it will contain an ImageSize object.
            ImageSize imageSize = shape.ImageData.ImageSize;

            // The ImageSize object contains read-only information about the image within the shape.
            Assert.That(imageSize.HeightPixels, Is.EqualTo(400));
            Assert.That(imageSize.WidthPixels, Is.EqualTo(400));

            const double delta = 0.05;
            Assert.That(imageSize.HorizontalResolution, Is.EqualTo(95.98d).Within(delta));
            Assert.That(imageSize.VerticalResolution, Is.EqualTo(95.98d).Within(delta));

            // We can base the size of the shape on the size of its image to avoid stretching the image.
            shape.Width = imageSize.WidthPoints * 2;
            shape.Height = imageSize.HeightPoints * 2;

            doc.Save(ArtifactsDir + "Drawing.ImageSize.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Drawing.ImageSize.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, shape);
            Assert.That(shape.Width, Is.EqualTo(600.0d));
            Assert.That(shape.Height, Is.EqualTo(600.0d));

            imageSize = shape.ImageData.ImageSize;

            Assert.That(imageSize.HeightPixels, Is.EqualTo(400));
            Assert.That(imageSize.WidthPixels, Is.EqualTo(400));
            Assert.That(imageSize.HorizontalResolution, Is.EqualTo(95.98d).Within(delta));
            Assert.That(imageSize.VerticalResolution, Is.EqualTo(95.98d).Within(delta));
        }
    }
}