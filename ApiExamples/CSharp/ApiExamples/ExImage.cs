// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Linq;
using System.Net;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;
#if NET462 || JAVA
using System.Drawing;
#elif NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    /// <summary>
    /// Mostly scenarios that deal with image shapes.
    /// </summary>
    [TestFixture]
    public class ExImage : ApiExampleBase
    {
        [Test]
        public void FromFile()
        {
            //ExStart
            //ExFor:Shape.#ctor(DocumentBase,ShapeType)
            //ExFor:ShapeType
            //ExSummary:Shows how to insert a shape with an image from the local file system into a document.
            Document doc = new Document();

            // The "Shape" class's public constructor will create a shape with "ShapeMarkupLanguage.Vml" markup type.
            // If you need to create a shape of a non-primitive type, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
            // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
            // please use DocumentBuilder.InsertShape.
            Shape shape = new Shape(doc, ShapeType.Image);
            shape.ImageData.SetImage(ImageDir + "Windows MetaFile.wmf");
            shape.Width = 100;
            shape.Height = 100;

            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            doc.Save(ArtifactsDir + "Image.FromFile.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.FromFile.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, shape);
            Assert.AreEqual(100.0d, shape.Height);
            Assert.AreEqual(100.0d, shape.Width);
        }

        [Test]
        public void FromUrl()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExSummary:Shows how to insert a shape with an image into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two locations where the document builder's "InsertShape" method
            // can source the image that the shape will display.
            // 1 -  Pass a local file system filename of an image file:
            builder.Write("Image from local file: ");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.Writeln();

            // 2 -  Pass a URL which points to an image.
            builder.Write("Image from a URL: ");
            builder.InsertImage(AsposeLogoUrl);
            builder.Writeln();

            doc.Save(ArtifactsDir + "Image.FromUrl.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.FromUrl.docx");
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            Assert.AreEqual(2, shapes.Count);
            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, (Shape)shapes[0]);
            TestUtil.VerifyImageInShape(320, 320, ImageType.Png, (Shape)shapes[1]);
        }

        [Test]
        public void FromStream()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream)
            //ExSummary:Shows how to insert a shape with an image from a stream into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (Stream stream = File.OpenRead(ImageDir + "Logo.jpg"))
            {
                builder.Write("Image from stream: ");
                builder.InsertImage(stream);
            }

            doc.Save(ArtifactsDir + "Image.FromStream.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.FromStream.docx");

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, (Shape)doc.GetChildNodes(NodeType.Shape, true)[0]);
        }

        #if NET462 || JAVA
        [Test, Category("SkipMono")]
        public void FromImage()
        {
            DocumentBuilder builder = new DocumentBuilder();

            using (Image rasterImage = Image.FromFile(ImageDir + "Logo.jpg"))
            {
                builder.Write("Raster image: ");
                builder.InsertImage(rasterImage);
                builder.Writeln();
            }

            using (Image metafile = Image.FromFile(ImageDir + "Windows MetaFile.wmf"))
            {
                builder.Write("Metafile: ");
                builder.InsertImage(metafile);
                builder.Writeln();
            }

            builder.Document.Save(ArtifactsDir + "Image.FromImage.docx");
        }
#elif NETCOREAPP2_1 || __MOBILE__
        [Test]
        [Category("SkipMono")]
        public void FromImageNetStandard2()
        {
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a raster image
            using (SKBitmap rasterImage = SKBitmap.Decode(ImageDir + "Logo.jpg"))
            {
                builder.Write("Raster image: ");
                builder.InsertImage(rasterImage);
                builder.Writeln();
            }

            builder.Document.Save(ArtifactsDir + "Image.FromImage.docx");
        }
#endif

        [Test]
        public void CreateFloatingPageCenter()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExFor:Shape
            //ExFor:ShapeBase
            //ExFor:ShapeBase.WrapType
            //ExFor:ShapeBase.BehindText
            //ExFor:ShapeBase.RelativeHorizontalPosition
            //ExFor:ShapeBase.RelativeVerticalPosition
            //ExFor:ShapeBase.HorizontalAlignment
            //ExFor:ShapeBase.VerticalAlignment
            //ExFor:WrapType
            //ExFor:RelativeHorizontalPosition
            //ExFor:RelativeVerticalPosition
            //ExFor:HorizontalAlignment
            //ExFor:VerticalAlignment
            //ExSummary:Shows how to insert a floating image to the center of a page.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a floating image that will appear behind the overlapping text and align it to the page's center.
            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");
            shape.WrapType = WrapType.None;
            shape.BehindText = true;
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            shape.HorizontalAlignment = HorizontalAlignment.Center;
            shape.VerticalAlignment = VerticalAlignment.Center;

            doc.Save(ArtifactsDir + "Image.CreateFloatingPageCenter.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.CreateFloatingPageCenter.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, shape);
            Assert.AreEqual(WrapType.None, shape.WrapType);
            Assert.True(shape.BehindText);
            Assert.AreEqual(RelativeHorizontalPosition.Page, shape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Page, shape.RelativeVerticalPosition);
            Assert.AreEqual(HorizontalAlignment.Center, shape.HorizontalAlignment);
            Assert.AreEqual(VerticalAlignment.Center, shape.VerticalAlignment);
        }

        [Test]
        public void CreateFloatingPositionSize()
        {
            //ExStart
            //ExFor:ShapeBase.Left
            //ExFor:ShapeBase.Right
            //ExFor:ShapeBase.Top
            //ExFor:ShapeBase.Bottom
            //ExFor:ShapeBase.Width
            //ExFor:ShapeBase.Height
            //ExFor:DocumentBuilder.CurrentSection
            //ExFor:PageSetup.PageWidth
            //ExSummary:Shows how to insert a floating image, and specify its position and size.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");
            shape.WrapType = WrapType.None;

            // Configure the shape's "RelativeHorizontalPosition" property to treat the value of the "Left" property
            // as the shape's horizontal distance, in points, from the left side of the page. 
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;

            // Set the shape's horizontal distance from the left side of the page to 100.
            shape.Left = 100;

            // Use the "RelativeVerticalPosition" property in a similar way to position the shape 80pt below the top of the page.
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            shape.Top = 80;

            // Set the shape's height, which will automatically scale the width to preserve dimensions.
            shape.Height = 125;

            Assert.AreEqual(125.0d, shape.Width);

            // The "Bottom" and "Right" properties contain the bottom and right edges of the image.
            Assert.AreEqual(shape.Top + shape.Height, shape.Bottom);
            Assert.AreEqual(shape.Left + shape.Width, shape.Right);

            doc.Save(ArtifactsDir + "Image.CreateFloatingPositionSize.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.CreateFloatingPositionSize.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, shape);
            Assert.AreEqual(WrapType.None, shape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Page, shape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Page, shape.RelativeVerticalPosition);
            Assert.AreEqual(100.0d, shape.Left);
            Assert.AreEqual(80.0d, shape.Top);
            Assert.AreEqual(125.0d, shape.Height);
            Assert.AreEqual(125.0d, shape.Width);
            Assert.AreEqual(shape.Top + shape.Height, shape.Bottom);
            Assert.AreEqual(shape.Left + shape.Width, shape.Right);
        }

        [Test]
        public void InsertImageWithHyperlink()
        {
            //ExStart
            //ExFor:ShapeBase.HRef
            //ExFor:ShapeBase.ScreenTip
            //ExFor:ShapeBase.Target
            //ExSummary:Shows how to insert a shape which contains an image, and is also a hyperlink.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");
            shape.HRef = "https://forum.aspose.com/";
            shape.Target = "New Window";
            shape.ScreenTip = "Aspose.Words Support Forums";

            // Ctrl + left-clicking the shape in Microsoft Word will open a new web browser window
            // and take us to the hyperlink in the "HRef" property.
            doc.Save(ArtifactsDir + "Image.InsertImageWithHyperlink.docx");
            //ExEnd
            
            doc = new Document(ArtifactsDir + "Image.InsertImageWithHyperlink.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, shape.HRef);
            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, shape);
            Assert.AreEqual("New Window", shape.Target);
            Assert.AreEqual("Aspose.Words Support Forums", shape.ScreenTip);
        }

        [Test]
        public void CreateLinkedImage()
        {
            //ExStart
            //ExFor:Shape.ImageData
            //ExFor:ImageData
            //ExFor:ImageData.SourceFullName
            //ExFor:ImageData.SetImage(String)
            //ExFor:DocumentBuilder.InsertNode
            //ExSummary:Shows how to insert a linked image into a document. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string imageFileName = ImageDir + "Windows MetaFile.wmf";

            // Below are two ways of applying an image to a shape so that it can display it.
            // 1 -  Set the shape to contain the image.
            Shape shape = new Shape(builder.Document, ShapeType.Image);
            shape.WrapType = WrapType.Inline;
            shape.ImageData.SetImage(imageFileName);

            builder.InsertNode(shape);

            doc.Save(ArtifactsDir + "Image.CreateLinkedImage.Embedded.docx");

            // Every image that we store in shape will increase the size of our document.
            Assert.True(70000 < new FileInfo(ArtifactsDir + "Image.CreateLinkedImage.Embedded.docx").Length);

            doc.FirstSection.Body.FirstParagraph.RemoveAllChildren();

            // 2 -  Set the shape to link to an image file in the local file system.
            shape = new Shape(builder.Document, ShapeType.Image);
            shape.WrapType = WrapType.Inline;
            shape.ImageData.SourceFullName = imageFileName;

            builder.InsertNode(shape);
            doc.Save(ArtifactsDir + "Image.CreateLinkedImage.Linked.docx");

            // Linking to images will save space and result in a smaller document.
            // However, the document can only display the image correctly while
            // the image file is present at the location that the shape's "SourceFullName" property points to.
            Assert.True(10000 > new FileInfo(ArtifactsDir + "Image.CreateLinkedImage.Linked.docx").Length);
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.CreateLinkedImage.Embedded.docx");

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, shape);
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            Assert.AreEqual(string.Empty, shape.ImageData.SourceFullName.Replace("%20", " "));

            doc = new Document(ArtifactsDir + "Image.CreateLinkedImage.Linked.docx");

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(0, 0, ImageType.Wmf, shape);
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            Assert.AreEqual(imageFileName, shape.ImageData.SourceFullName.Replace("%20", " "));
        }

        [Test]
        public void DeleteAllImages()
        {
            //ExStart
            //ExFor:Shape.HasImage
            //ExFor:Node.Remove
            //ExSummary:Shows how to delete all shapes with images from a document.
            Document doc = new Document(MyDir + "Images.docx");
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            Assert.AreEqual(9, shapes.OfType<Shape>().Count(s => s.HasImage));

            foreach (Shape shape in shapes.OfType<Shape>())
                if (shape.HasImage) 
                    shape.Remove();

            Assert.AreEqual(0, shapes.OfType<Shape>().Count(s => s.HasImage));
            //ExEnd
        }

        [Test]
        public void DeleteAllImagesPreOrder()
        {
            //ExStart
            //ExFor:Node.NextPreOrder(Node)
            //ExFor:Node.PreviousPreOrder(Node)
            //ExSummary:Shows how to traverse the document's node tree using the pre-order traversal algorithm, and delete any encountered shape with an image.
            Document doc = new Document(MyDir + "Images.docx");

            Assert.AreEqual(9, 
                doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().Count(s => s.HasImage));

            Node curNode = doc;
            while (curNode != null)
            {
                Node nextNode = curNode.NextPreOrder(doc);

                if (curNode.PreviousPreOrder(doc) != null && nextNode != null)
                    Assert.AreEqual(curNode, nextNode.PreviousPreOrder(doc));

                if (curNode.NodeType == NodeType.Shape && ((Shape)curNode).HasImage)
                    curNode.Remove();
                
                curNode = nextNode;
            }

            Assert.AreEqual(0,
                doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().Count(s => s.HasImage));
            //ExEnd
        }

        [Test]
        public void ScaleImage()
        {
            //ExStart
            //ExFor:ImageData.ImageSize
            //ExFor:ImageSize
            //ExFor:ImageSize.WidthPoints
            //ExFor:ImageSize.HeightPoints
            //ExFor:ShapeBase.Width
            //ExFor:ShapeBase.Height
            //ExSummary:Shows how to resize a shape with an image.
#if NET462 || JAVA
            Image image = Image.FromFile(ImageDir + "Logo.jpg");

            Assert.AreEqual(400, image.Size.Width);
            Assert.AreEqual(400, image.Size.Height);
#elif NETCOREAPP2_1
            SKBitmap image = SKBitmap.Decode(ImageDir + "Logo.jpg");

            Assert.AreEqual(400, image.Width);
            Assert.AreEqual(400, image.Height);
#endif

            // When we insert an image using the "InsertImage" method, the builder scales the shape that displays the image so that,
            // when we view the document using 100% zoom in Microsoft Word, the shape displays the image in its actual size.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");

            // A 400x400 image will create an ImageData object with an image size of 300x300pt.
            ImageSize imageSize = shape.ImageData.ImageSize;

            Assert.AreEqual(300.0d, imageSize.WidthPoints);
            Assert.AreEqual(300.0d, imageSize.HeightPoints);

            // If a shape's dimensions match the image data's dimensions,
            // then the shape is displaying the image in its original size.
            Assert.AreEqual(300.0d, shape.Width);
            Assert.AreEqual(300.0d, shape.Height);

            // Reduce the overall size of the shape by 50%. 
            shape.Width *= 0.5;

            // Scaling factors apply to both the width and the height at the same time to preserve the shape's proportions. 
            Assert.AreEqual(150.0d, shape.Width);
            Assert.AreEqual(150.0d, shape.Height);

            // When we resize the shape, the size of the image data remains the same.
            Assert.AreEqual(300.0d, imageSize.WidthPoints);
            Assert.AreEqual(300.0d, imageSize.HeightPoints);

            // We can reference the image data dimensions to apply a scaling based on the size of the image.
            shape.Width = imageSize.WidthPoints * 1.1;

            Assert.AreEqual(330.0d, shape.Width);
            Assert.AreEqual(330.0d, shape.Height);

            doc.Save(ArtifactsDir + "Image.ScaleImage.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.ScaleImage.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(330.0d, shape.Width);
            Assert.AreEqual(330.0d, shape.Height);

            imageSize = shape.ImageData.ImageSize;

            Assert.AreEqual(300.0d, imageSize.WidthPoints);
            Assert.AreEqual(300.0d, imageSize.HeightPoints);
        }
    }
}