using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;
using Shape = Aspose.Words.Drawing.Shape;
#if !(NETSTANDARD2_0 || __MOBILE__)
using System.Net;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDrawing : ApiExampleBase
    {
#if !(NETSTANDARD2_0 || __MOBILE__)
        [Test]
        public void DrawingVariousShapes()
        {
            //ExStart
            //ExFor:Drawing.ArrowLength
            //ExFor:Drawing.ArrowType
            //ExFor:Drawing.ArrowWidth
            //ExFor:Drawing.DashStyle
            //ExFor:Drawing.EndCap
            //ExFor:Drawing.Fill.Color
            //ExFor:Drawing.Fill.ImageBytes
            //ExFor:Drawing.Fill.On
            //ExFor:Drawing.JoinStyle
            //ExFor:Stroke.Color
            //ExFor:Stroke.StartArrowLength
            //ExFor:Stroke.StartArrowType
            //ExFor:Stroke.StartArrowWidth
            //ExFor:Stroke.DashStyle
            //ExFor:Stroke.EndArrowType
            //ExFor:Stroke.EndCap
            //ExSummary:Shows to create a variety of shapes
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Draw a dotted horizontal red line with an arrow on the left end and a diamond on the other
            Shape arrow = new Shape(doc, ShapeType.Line);
            arrow.Width = 200;
            arrow.Stroke.Color = Color.Red;
            arrow.Stroke.StartArrowType = ArrowType.Arrow;
            arrow.Stroke.StartArrowLength = ArrowLength.Long;
            arrow.Stroke.StartArrowWidth = ArrowWidth.Wide;
            arrow.Stroke.EndArrowType = ArrowType.Diamond;
            arrow.Stroke.DashStyle = DashStyle.Dash;

            Assert.AreEqual(JoinStyle.Miter, arrow.Stroke.JoinStyle);

            builder.InsertNode(arrow);

            // Draw a thick black diagonal line with rounded ends
            Shape line = new Shape(doc, ShapeType.Line);
            line.Top = 40;
            line.Width = 200;
            line.Height = 20;
            line.StrokeWeight = 5.0;
            line.Stroke.EndCap = EndCap.Round;

            builder.InsertNode(line);

            // Draw an arrow with a green fill
            Shape filledInArrow = new Shape(doc, ShapeType.Arrow);
            filledInArrow.Width = 200;
            filledInArrow.Height = 40;
            filledInArrow.Top = 100;
            filledInArrow.Fill.Color = Color.Green;
            filledInArrow.Fill.On = true;

            builder.InsertNode(filledInArrow);

            // Draw an arrow filled in with the Aspose logo and flip its orientation
            Shape filledInArrowImg = new Shape(doc, ShapeType.Arrow);
            filledInArrowImg.Width = 200;
            filledInArrowImg.Height = 40;
            filledInArrowImg.Top = 160;
            filledInArrowImg.FlipOrientation = FlipOrientation.Both;

            using (WebClient webClient = new WebClient())
            {
                byte[] imageBytes = webClient.DownloadData("http://www.aspose.com/images/aspose-logo.gif");

                using (System.IO.MemoryStream stream = new System.IO.MemoryStream(imageBytes))
                {
                    Image image = Image.FromStream(stream);
                    // When we flipped the orientation of our arrow, the image content was flipped too
                    // If we want it to be displayed the right side up, we have to reverse the arrow flip on the image
                    image.RotateFlip(RotateFlipType.RotateNoneFlipXY);

                    filledInArrowImg.ImageData.SetImage(image);
                    builder.InsertNode(filledInArrowImg);

                    filledInArrowImg.Stroke.JoinStyle = JoinStyle.Round;
                }
            }

            doc.Save(ArtifactsDir + "Drawing.VariousShapes.docx");
            //ExEnd
        }
#endif

        //ExStart
        //ExFor:DocumentVisitor.VisitShapeEnd(Shape)
        //ExFor:DocumentVisitor.VisitShapeStart(Shape)
        //ExFor:DocumentVisitor.VisitGroupShapeEnd(GroupShape)
        //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
        //ExFor:Drawing.GroupShape
        //ExFor:Drawing.GroupShape.#ctor(DocumentBase)
        //ExFor:Drawing.GroupShape.#ctor(DocumentBase,Drawing.ShapeMarkupLanguage)
        //ExFor:Drawing.GroupShape.Accept(DocumentVisitor)
        //ExSummary:Shows how to create a group of shapes, and let it accept a visitor
        [Test] //ExSkip
        public void GroupOfShapes()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
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

            builder.InsertNode(group);

            ShapeInfoPrinter printer = new ShapeInfoPrinter();
            group.Accept(printer);

            Console.WriteLine(printer.GetText());
        }

        /// <summary>
        /// Visitor that prints shape group contents information to the console.
        /// </summary>
        public class ShapeInfoPrinter : DocumentVisitor
        {
            public ShapeInfoPrinter()
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
                mBuilder.AppendLine("\t\tFill color: " + shape.Fill.Color);
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

#if !(NETSTANDARD2_0 || __MOBILE__)
        [Test]
        public void TypeOfImage()
        {
            //ExStart
            //ExFor:Drawing.ImageType
            //ExSummary:Shows how to add an image to a shape and check its type
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (WebClient webClient = new WebClient())
            {
                byte[] imageBytes = webClient.DownloadData("http://www.aspose.com/images/aspose-logo.gif");

                using (System.IO.MemoryStream stream = new System.IO.MemoryStream(imageBytes))
                {
                    Image image = Image.FromStream(stream);

                    // The image started off as an animated .gif but it gets converted to a .png since there cannot be animated images in documents
                    Shape imgShape = builder.InsertImage(image);
                    Assert.AreEqual(ImageType.Png, imgShape.ImageData.ImageType);
                }
            }

            //ExEnd
        }
#endif

        [Test]
        public void TextBoxTextLayout()
        {
            //ExStart
            //ExFor:Drawing.LayoutFlow
            //ExSummary:Shows how to add text to a textbox and change its orientation
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
        }

        [Test]
        public void GetDataFromImage()
        {
            //ExStart
            //ExFor:ImageData.ImageBytes
            //ExFor:ImageData.ToByteArray
            //ExFor:ImageData.ToStream
            //ExSummary:Shows how to access raw image data in a shape's ImageData object.
            Document imgSourceDoc = new Document(MyDir + "Image.SampleImages.doc");

            // Images are stored as shapes
            // Get into the document's shape collection to verify that it contains 6 images
            List<Shape> shapes = imgSourceDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(6, shapes.Count);

            // ToByteArray() returns the value of the ImageBytes property
            Assert.AreEqual(shapes[0].ImageData.ImageBytes, shapes[0].ImageData.ToByteArray());

            // Put the shape's image data into a stream
            // Then, put the image data from that stream into another stream which uses that data to create an image
            using (Stream imgStream = shapes[0].ImageData.ToStream())
            {
                FileStream outStream = new FileStream(ArtifactsDir + "MyImg.png", FileMode.CreateNew);
                imgStream.CopyTo(outStream);
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
            //ExSummary:Shows how to save all the images from a document to the file system.
            Document imgSourceDoc = new Document(MyDir + "Image.SampleImages.doc");

            // Images are stored as shapes
            // Get into the document's shape collection to verify that it contains 6 images
            List<Shape> shapes = imgSourceDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
            Assert.AreEqual(6, shapes.Count);

            // We will use an ImageFormatConverter to determine an image's file extension
            ImageFormatConverter formatConverter = new ImageFormatConverter();

            // Go over all of the document's shapes
            // If a shape contains image data, save the image in the local file system
            for (int i = 0; i < shapes.Count; i++)
            {
                ImageData imageData = shapes[i].ImageData;

                if (imageData.HasImage)
                {
                    ImageFormat format = imageData.ToImage().RawFormat;
                    string fileExtension = formatConverter.ConvertToString(format);

                    using (FileStream fileStream = File.Create(ArtifactsDir + $"Image from shape {i}.{fileExtension}"))
                    {
                        imageData.Save(fileStream);
                    }
                }
            }
            //ExEnd
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
            //ExSummary:Shows how to edit images using the ImageData attribute.
            // Open a document that contains images
            Document imgSourceDoc = new Document(MyDir + "Image.SampleImages.doc");

            Shape sourceShape = (Shape)imgSourceDoc.GetChildNodes(NodeType.Shape, true)[0];
            
            Document dstDoc = new Document();

            // Import a shape from the source document and append it to the first paragraph, effectively cloning it
            Shape importedShape = (Shape)dstDoc.ImportNode(sourceShape, true);
            dstDoc.FirstSection.Body.FirstParagraph.AppendChild(importedShape);

            // Get the ImageData of the imported shape
            ImageData imageData = importedShape.ImageData;
            imageData.Title = "Imported Image";

            // If an image appears to have no borders, its ImageData object will still have them, but in an unspecified color
            Assert.AreEqual(4, imageData.Borders.Count);
            Assert.AreEqual(Color.Empty, imageData.Borders[0].Color);

            Assert.True(imageData.HasImage);

            // This image is not linked to a shape or to an image in the file system
            Assert.False(imageData.IsLink);
            Assert.False(imageData.IsLinkOnly);

            // Brightness and contrast are defined on a 0-1 scale, with 0.5 being the default value
            imageData.Brightness = 0.8d;
            imageData.Contrast = 1.0d;

            // Our image will have a lot of white now that we've changed the brightness and contrast like that
            // We can treat white as transparent with the following attribute
            imageData.ChromaKey = Color.White;

            // Import the source shape again, set it to black and white
            importedShape = (Shape)dstDoc.ImportNode(sourceShape, true);
            dstDoc.FirstSection.Body.FirstParagraph.AppendChild(importedShape);

            importedShape.ImageData.GrayScale = true;

            // Import the source shape again to create a third image, and set it to BiLevel
            // Unlike greyscale, which preserves the brightness of the original colors,
            // BiLevel sets every pixel to either black or white, whichever is closer to the original color
            importedShape = (Shape)dstDoc.ImportNode(sourceShape, true);
            dstDoc.FirstSection.Body.FirstParagraph.AppendChild(importedShape);

            importedShape.ImageData.BiLevel = true;

            // Cropping is determined on a 0-1 scale
            // Cropping a side by 0.3 will crop 30% of the image out at that side
            importedShape.ImageData.CropBottom = 0.3d;
            importedShape.ImageData.CropLeft = 0.3d;
            importedShape.ImageData.CropTop = 0.3d;
            importedShape.ImageData.CropRight = 0.3d;

            dstDoc.Save(ArtifactsDir + "ImageData.EditedImages.docx");
            //ExEnd
        }

        [Test]
        public void ImportImage()
        {
            //ExStart
            //ExFor:ImageData.SetImage(Image)
            //ExFor:ImageData.SetImage(Stream)
            //ExSummary:Shows two ways of importing images from the local file system into a document.
            Document doc = new Document();

            // We can get an image from a file, set it as the image of a shape and append it to a paragraph
            Image srcImage = Image.FromFile(ImageDir + "Aspose.Words.gif");

            Shape imgShape = new Shape(doc, ShapeType.Image);
            doc.FirstSection.Body.FirstParagraph.AppendChild(imgShape);
            imgShape.ImageData.SetImage(srcImage);
            srcImage.Dispose();

            // We can also open an image file using a stream and set its contents as a shape's image 
            using (Stream stream = new FileStream(ImageDir + "Aspose.Words.gif", FileMode.Open, FileAccess.Read))
            {
                imgShape = new Shape(doc, ShapeType.Image);
                doc.FirstSection.Body.FirstParagraph.AppendChild(imgShape);
                imgShape.ImageData.SetImage(stream);
                imgShape.Left = 150.0f;
            }

            doc.Save(ArtifactsDir + "ImageData.ImportedImage.docx");
            //ExEnd
        }
    }
}