// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
#if NET462 || JAVA
using System.Drawing.Imaging;
using System.Net;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDrawing : ApiExampleBase
    {
        #if NET462 || JAVA
        [Test]
        public void VariousShapes()
        {
            //ExStart
            //ExFor:Drawing.ArrowLength
            //ExFor:Drawing.ArrowType
            //ExFor:Drawing.ArrowWidth
            //ExFor:Drawing.DashStyle
            //ExFor:Drawing.EndCap
            //ExFor:Drawing.Fill.ForeColor
            //ExFor:Drawing.Fill.ImageBytes
            //ExFor:Drawing.Fill.Visible
            //ExFor:Drawing.JoinStyle
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

            Assert.AreEqual(JoinStyle.Miter, arrow.Stroke.JoinStyle);

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

            using (WebClient webClient = new WebClient())
            {
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
            }

            doc.Save(ArtifactsDir + "Drawing.VariousShapes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Drawing.VariousShapes.docx");

            Assert.AreEqual(4, doc.GetChildNodes(NodeType.Shape, true).Count);

            arrow = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(ShapeType.Line, arrow.ShapeType);
            Assert.AreEqual(200.0d, arrow.Width);
            Assert.AreEqual(Color.Red.ToArgb(), arrow.Stroke.Color.ToArgb());
            Assert.AreEqual(ArrowType.Arrow, arrow.Stroke.StartArrowType);
            Assert.AreEqual(ArrowLength.Long, arrow.Stroke.StartArrowLength);
            Assert.AreEqual(ArrowWidth.Wide, arrow.Stroke.StartArrowWidth);
            Assert.AreEqual(ArrowType.Diamond, arrow.Stroke.EndArrowType);
            Assert.AreEqual(ArrowLength.Long, arrow.Stroke.EndArrowLength);
            Assert.AreEqual(ArrowWidth.Wide, arrow.Stroke.EndArrowWidth);
            Assert.AreEqual(DashStyle.Dash, arrow.Stroke.DashStyle);
            Assert.AreEqual(0.5d, arrow.Stroke.Opacity);

            line = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(ShapeType.Line, line.ShapeType);
            Assert.AreEqual(40.0d, line.Top);
            Assert.AreEqual(200.0d, line.Width);
            Assert.AreEqual(20.0d, line.Height);
            Assert.AreEqual(5.0d, line.StrokeWeight);
            Assert.AreEqual(EndCap.Round, line.Stroke.EndCap);

            filledInArrow = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.AreEqual(ShapeType.Arrow, filledInArrow.ShapeType);
            Assert.AreEqual(200.0d, filledInArrow.Width);
            Assert.AreEqual(40.0d, filledInArrow.Height);
            Assert.AreEqual(100.0d, filledInArrow.Top);
            Assert.AreEqual(Color.Green.ToArgb(), filledInArrow.Fill.ForeColor.ToArgb());
            Assert.True(filledInArrow.Fill.Visible);

            filledInArrowImg = (Shape)doc.GetChild(NodeType.Shape, 3, true);

            Assert.AreEqual(ShapeType.Arrow, filledInArrowImg.ShapeType);
            Assert.AreEqual(200.0d, filledInArrowImg.Width);
            Assert.AreEqual(40.0d, filledInArrowImg.Height);
            Assert.AreEqual(160.0d, filledInArrowImg.Top);
            Assert.AreEqual(FlipOrientation.Both, filledInArrowImg.FlipOrientation);
        }

        [Test]
        public void TypeOfImage()
        {
            //ExStart
            //ExFor:Drawing.ImageType
            //ExSummary:Shows how to add an image to a shape and check its type.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (WebClient webClient = new WebClient())
            {
                byte[] imageBytes = File.ReadAllBytes(ImageDir + "Logo.jpg");

                using (MemoryStream stream = new MemoryStream(imageBytes))
                {
                    Image image = Image.FromStream(stream);

                    // The image in the URL is a .gif. Inserting it into a document converts it into a .png.
                    Shape imgShape = builder.InsertImage(image);
                    Assert.AreEqual(ImageType.Jpeg, imgShape.ImageData.ImageType);
                }
            }
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
            IEnumerable<Shape> shapesWithImages = 
                imgSourceDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage);

            // Go through each shape and save its image.
            ImageFormatConverter formatConverter = new ImageFormatConverter();

            using (IEnumerator<Shape> enumerator = shapesWithImages.GetEnumerator())
            {
                int shapeIndex = 0;

                while (enumerator.MoveNext())
                {
                    ImageData imageData = enumerator.Current.ImageData;
                    ImageFormat format = imageData.ToImage().RawFormat;
                    string fileExtension = formatConverter.ConvertToString(format);

                    using (FileStream fileStream = File.Create(ArtifactsDir + $"Drawing.SaveAllImages.{++shapeIndex}.{fileExtension}"))
                        imageData.Save(fileStream);
                }
            }
            //ExEnd

            string[] imageFileNames = Directory.GetFiles(ArtifactsDir).Where(s => s.StartsWith(ArtifactsDir + "Drawing.SaveAllImages.")).OrderBy(s => s).ToArray();
            List<FileInfo> fileInfos = imageFileNames.Select(s => new FileInfo(s)).ToList();
            
            TestUtil.VerifyImage(2467, 1500, fileInfos[0].FullName);
            Assert.AreEqual(".Jpeg", fileInfos[0].Extension);
            TestUtil.VerifyImage(400, 400, fileInfos[1].FullName);
            Assert.AreEqual(".Png", fileInfos[1].Extension);
            TestUtil.VerifyImage(382, 138, fileInfos[2].FullName);
            Assert.AreEqual(".Emf", fileInfos[2].Extension);
            TestUtil.VerifyImage(1600, 1600, fileInfos[3].FullName);
            Assert.AreEqual(".Wmf", fileInfos[3].Extension);
            TestUtil.VerifyImage(534, 534, fileInfos[4].FullName);
            Assert.AreEqual(".Emf", fileInfos[4].Extension);
            TestUtil.VerifyImage(1260, 660, fileInfos[5].FullName);
            Assert.AreEqual(".Jpeg", fileInfos[5].Extension);
            TestUtil.VerifyImage(1125, 1500, fileInfos[6].FullName);
            Assert.AreEqual(".Jpeg", fileInfos[6].Extension);
            TestUtil.VerifyImage(1027, 1500, fileInfos[7].FullName);
            Assert.AreEqual(".Jpeg", fileInfos[7].Extension);
            TestUtil.VerifyImage(1200, 1500, fileInfos[8].FullName);
            Assert.AreEqual(".Jpeg", fileInfos[8].Extension);
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

            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Shape, true).Count);

            imgShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imgShape);
            Assert.AreEqual(0.0d, imgShape.Left);
            Assert.AreEqual(0.0d, imgShape.Top);
            Assert.AreEqual(300.0d, imgShape.Height);
            Assert.AreEqual(300.0d, imgShape.Width);
            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imgShape);

            imgShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imgShape);
            Assert.AreEqual(150.0d, imgShape.Left);
            Assert.AreEqual(0.0d, imgShape.Top);
            Assert.AreEqual(300.0d, imgShape.Height);
            Assert.AreEqual(300.0d, imgShape.Width);
        }
#endif

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
            Assert.AreEqual(Color.FromArgb(255, 128, 0, 0), stroke.Color);
            Assert.AreEqual(Color.FromArgb(255, 255, 255, 0), stroke.Color2);

            Assert.NotNull(stroke.ImageBytes);
            File.WriteAllBytes(ArtifactsDir + "Drawing.StrokePattern.png", stroke.ImageBytes);
            //ExEnd

            TestUtil.VerifyImage(8, 8, ArtifactsDir + "Drawing.StrokePattern.png");
        }

        //ExStart
        //ExFor:DocumentVisitor.VisitShapeEnd(Shape)
        //ExFor:DocumentVisitor.VisitShapeStart(Shape)
        //ExFor:DocumentVisitor.VisitGroupShapeEnd(GroupShape)
        //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
        //ExFor:Drawing.GroupShape
        //ExFor:Drawing.GroupShape.#ctor(DocumentBase)
        //ExFor:Drawing.GroupShape.Accept(DocumentVisitor)
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

            Assert.True(group.IsGroup);

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

            Assert.AreEqual(2, shapes.ChildNodes.Count);

            Shape shape = (Shape)shapes.ChildNodes[0];

            Assert.AreEqual(ShapeType.Balloon, shape.ShapeType);
            Assert.AreEqual(200.0d, shape.Width);
            Assert.AreEqual(200.0d, shape.Height);
            Assert.AreEqual(Color.Red.ToArgb(), shape.StrokeColor.ToArgb());

            shape = (Shape)shapes.ChildNodes[1];

            Assert.AreEqual(ShapeType.Cube, shape.ShapeType);
            Assert.AreEqual(100.0d, shape.Width);
            Assert.AreEqual(100.0d, shape.Height);
            Assert.AreEqual(Color.Blue.ToArgb(), shape.StrokeColor.ToArgb());
        }

        [Test]
        public void TextBox()
        {
            //ExStart
            //ExFor:Drawing.LayoutFlow
            //ExSummary:Shows how to add text to a text box, and change its orientation
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textbox = new Shape(doc, ShapeType.TextBox)
            {
                Width = 100, 
                Height = 100,
                TextBox = { LayoutFlow = LayoutFlow.BottomToTop }
            };
            
            textbox.AppendChild(new Paragraph(doc));
            builder.InsertNode(textbox);

            builder.MoveTo(textbox.FirstParagraph);
            builder.Write("This text is flipped 90 degrees to the left.");
            
            doc.Save(ArtifactsDir + "Drawing.TextBox.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Drawing.TextBox.docx");
            textbox = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(ShapeType.TextBox, textbox.ShapeType);
            Assert.AreEqual(100.0d, textbox.Width);
            Assert.AreEqual(100.0d, textbox.Height);
            Assert.AreEqual(LayoutFlow.BottomToTop, textbox.TextBox.LayoutFlow);
            Assert.AreEqual("This text is flipped 90 degrees to the left.", textbox.GetText().Trim());
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
            Assert.AreEqual(10, imgSourceDoc.GetChildNodes(NodeType.Shape, true).Count); //ExSkip

            Shape imgShape = (Shape)imgSourceDoc.GetChild(NodeType.Shape, 0, true);

            Assert.True(imgShape.HasImage);

            // ToByteArray() returns the array stored in the ImageBytes property.
            Assert.AreEqual(imgShape.ImageData.ImageBytes, imgShape.ImageData.ToByteArray());

            // Save the shape's image data to an image file in the local file system.
            using (Stream imgStream = imgShape.ImageData.ToStream())
            {
                using (FileStream outStream = new FileStream(ArtifactsDir + "Drawing.GetDataFromImage.png", FileMode.Create, FileAccess.ReadWrite))
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

            Assert.True(imageData.HasImage);

            // If an image has no borders, its ImageData object will define the border color as empty.
            Assert.AreEqual(4, imageData.Borders.Count);
            Assert.AreEqual(Color.Empty, imageData.Borders[0].Color);

            // This image does not link to another shape or image file in the local file system.
            Assert.False(imageData.IsLink);
            Assert.False(imageData.IsLinkOnly);

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
            Assert.AreEqual("Imported Image", sourceShape.ImageData.Title);
            Assert.AreEqual(0.8d, sourceShape.ImageData.Brightness, 0.1d);
            Assert.AreEqual(1.0d, sourceShape.ImageData.Contrast, 0.1d);
            Assert.AreEqual(Color.White.ToArgb(), sourceShape.ImageData.ChromaKey.ToArgb());

            sourceShape = (Shape)imgSourceDoc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(2467, 1500, ImageType.Jpeg, sourceShape);
            Assert.True(sourceShape.ImageData.GrayScale);

            sourceShape = (Shape)imgSourceDoc.GetChild(NodeType.Shape, 2, true);

            TestUtil.VerifyImageInShape(2467, 1500, ImageType.Jpeg, sourceShape);
            Assert.True(sourceShape.ImageData.BiLevel);
            Assert.AreEqual(0.3d, sourceShape.ImageData.CropBottom, 0.1d);
            Assert.AreEqual(0.3d, sourceShape.ImageData.CropLeft, 0.1d);
            Assert.AreEqual(0.3d, sourceShape.ImageData.CropTop, 0.1d);
            Assert.AreEqual(0.3d, sourceShape.ImageData.CropRight, 0.1d);
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
            Assert.AreEqual(400, imageSize.HeightPixels);
            Assert.AreEqual(400, imageSize.WidthPixels);

			const double delta = 0.05;
            Assert.AreEqual(95.98d, imageSize.HorizontalResolution, delta);
            Assert.AreEqual(95.98d, imageSize.VerticalResolution, delta);

            // We can base the size of the shape on the size of its image to avoid stretching the image.
            shape.Width = imageSize.WidthPoints * 2;
            shape.Height = imageSize.HeightPoints * 2;

            doc.Save(ArtifactsDir + "Drawing.ImageSize.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Drawing.ImageSize.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, shape);
            Assert.AreEqual(600.0d, shape.Width);
            Assert.AreEqual(600.0d, shape.Height);

            imageSize = shape.ImageData.ImageSize;

            Assert.AreEqual(400, imageSize.HeightPixels);
            Assert.AreEqual(400, imageSize.WidthPixels);
            Assert.AreEqual(95.98d, imageSize.HorizontalResolution, delta);
            Assert.AreEqual(95.98d, imageSize.VerticalResolution, delta);
        }
    }
}