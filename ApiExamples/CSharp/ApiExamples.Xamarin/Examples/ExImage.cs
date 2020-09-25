// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections;
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
        public void CreateImageDirectly()
        {
            //ExStart
            //ExFor:Shape.#ctor(DocumentBase,ShapeType)
            //ExFor:ShapeType
            //ExSummary:Shows how to add a shape with an image to a document.
            Document doc = new Document();

            // Public constructor of "Shape" class creates shape with "ShapeMarkupLanguage.Vml" markup type
            // If you need to create non-primitive shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
            // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
            // please use DocumentBuilder.InsertShape
            Shape shape = new Shape(doc, ShapeType.Image);
            shape.ImageData.SetImage(ImageDir + "Windows MetaFile.wmf");
            shape.Width = 100;
            shape.Height = 100;

            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            doc.Save(ArtifactsDir + "Image.CreateImageDirectly.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.CreateImageDirectly.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, shape);
            Assert.AreEqual(100.0d, shape.Height);
            Assert.AreEqual(100.0d, shape.Width);
        }

        [Test]
        public void CreateFromUrl()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExSummary:Shows how to inserts an image from a URL.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Image from local file: ");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.Writeln();

            builder.Write("Image from a URL: ");
            builder.InsertImage(AsposeLogoUrl);
            builder.Writeln();

            doc.Save(ArtifactsDir + "Image.CreateFromUrl.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.CreateFromUrl.docx");
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            Assert.AreEqual(2, shapes.Count);
            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, (Shape)shapes[0]);
            TestUtil.VerifyImageInShape(320, 320, ImageType.Png, (Shape)shapes[1]);
        }

        [Test]
        public void CreateFromStream()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream)
            //ExSummary:Shows how to insert an image from a stream. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (Stream stream = File.OpenRead(ImageDir + "Logo.jpg"))
            {
                builder.Write("Image from stream: ");
                builder.InsertImage(stream);
            }

            doc.Save(ArtifactsDir + "Image.CreateFromStream.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.CreateFromStream.docx");

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, (Shape)doc.GetChildNodes(NodeType.Shape, true)[0]);
        }

        #if NET462 || JAVA
        [Test]
        [Category("SkipMono")]
        public void CreateFromImage()
        {
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a raster image
            using (Image rasterImage = Image.FromFile(ImageDir + "Logo.jpg"))
            {
                builder.Write("Raster image: ");
                builder.InsertImage(rasterImage);
                builder.Writeln();
            }

            // Aspose.Words allows to insert a metafile too
            using (Image metafile = Image.FromFile(ImageDir + "Windows MetaFile.wmf"))
            {
                builder.Write("Metafile: ");
                builder.InsertImage(metafile);
                builder.Writeln();
            }

            builder.Document.Save(ArtifactsDir + "Image.CreateFromImage.docx");
        }
        #elif NETCOREAPP2_1 || __MOBILE__
        [Test]
        [Category("SkipMono")]
        public void CreateFromImageNetStandard2()
        {
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a raster image
            // SKBitmap doesn't allow to insert a metafiles
            using (SKBitmap rasterImage = SKBitmap.Decode(ImageDir + "Logo.jpg"))
            {
                builder.Write("Raster image: ");
                builder.InsertImage(rasterImage);
                builder.Writeln();
            }

            builder.Document.Save(ArtifactsDir + "Image.CreateFromImage.docx");
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
            //ExSummary:Shows how to insert a floating image in the middle of a page.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default, the image is inline
            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");

            // Make the image float, put it behind text and center on the page
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
            //ExSummary:Shows how to insert a floating image and specify its position and size.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default, the image is inline
            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");

            // Make the image float, put it behind text and center on the page
            shape.WrapType = WrapType.None;

            // Make position relative to the page
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Set the shape's coordinates, from the top left corner of the page
            shape.Left = 100;
            shape.Top = 80;

            // Set the shape's height
            shape.Height = 125.0;

            // The width will be scaled to the height and the dimensions of the real image
            Assert.AreEqual(125.0, shape.Width);

            // The Bottom and Right members contain the locations of the bottom and right edges of the image
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
            //ExSummary:Shows how to insert an image with a hyperlink.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertImage(ImageDir + "Windows MetaFile.wmf");
            shape.HRef = "https://forum.aspose.com/";
            shape.Target = "New Window";
            shape.ScreenTip = "Aspose.Words Support Forums";

            doc.Save(ArtifactsDir + "Image.InsertImageWithHyperlink.docx");
            //ExEnd
            
            doc = new Document(ArtifactsDir + "Image.InsertImageWithHyperlink.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, shape.HRef);
            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, shape);
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

            builder.Write("Image linked, not stored in the document: ");

            Shape shape = new Shape(builder.Document, ShapeType.Image);
            shape.WrapType = WrapType.Inline;
            shape.ImageData.SourceFullName = imageFileName;

            builder.InsertNode(shape);
            builder.Writeln();

            builder.Write("Image linked and stored in the document: ");

            shape = new Shape(builder.Document, ShapeType.Image);
            shape.WrapType = WrapType.Inline;
            shape.ImageData.SourceFullName = imageFileName;
            shape.ImageData.SetImage(imageFileName);

            builder.InsertNode(shape);
            builder.Writeln();

            builder.Write("Image stored in the document, but not linked: ");

            shape = new Shape(builder.Document, ShapeType.Image);
            shape.WrapType = WrapType.Inline;
            shape.ImageData.SetImage(imageFileName);

            builder.InsertNode(shape);
            builder.Writeln();

            doc.Save(ArtifactsDir + "Image.CreateLinkedImage.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Image.CreateLinkedImage.docx");

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(0, 0, ImageType.Wmf, shape);
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            Assert.AreEqual(imageFileName, shape.ImageData.SourceFullName.Replace("%20", " "));

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, shape);
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            Assert.AreEqual(imageFileName, shape.ImageData.SourceFullName.Replace("%20", " "));

            shape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, shape);
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            Assert.AreEqual(string.Empty, shape.ImageData.SourceFullName.Replace("%20", " "));
        }

        [Test]
        public void DeleteAllImages()
        {
            //ExStart
            //ExFor:Shape.HasImage
            //ExFor:Node.Remove
            //ExSummary:Shows how to delete all images from a document.
            Document doc = new Document(MyDir + "Images.docx");
            Assert.AreEqual(10, doc.GetChildNodes(NodeType.Shape, true).Count);

            // Here we get all shapes from the document node, but you can do this for any smaller
            // node too, for example delete shapes from a single section or a paragraph
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            // We cannot delete shape nodes while we enumerate through the collection
            // One solution is to add nodes that we want to delete to a temporary array and delete afterwards
            ArrayList shapesToDelete = new ArrayList();

            // Several shape types can have an image including image shapes and OLE objects
            foreach (Shape shape in shapes.OfType<Shape>())
                if (shape.HasImage)
                    shapesToDelete.Add(shape);

            // Now we can delete shapes
            foreach (Shape shape in shapesToDelete)
                shape.Remove();

            // The only remaining shape doesn't have an image
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, true).Count);
            Assert.False(((Shape)doc.GetChild(NodeType.Shape, 0, true)).HasImage);
            //ExEnd
        }

        [Test]
        public void DeleteAllImagesPreOrder()
        {
            //ExStart
            //ExFor:Node.NextPreOrder(Node)
            //ExFor:Node.PreviousPreOrder(Node)
            //ExSummary:Shows how to delete all images from a document using pre-order tree traversal.
            Document doc = new Document(MyDir + "Images.docx");
            Assert.AreEqual(10, doc.GetChildNodes(NodeType.Shape, true).Count);

            Node curNode = doc;
            while (curNode != null)
            {
                Node nextNode = curNode.NextPreOrder(doc);

                if (curNode.PreviousPreOrder(doc) != null && nextNode != null)
                    Assert.AreEqual(curNode, nextNode.PreviousPreOrder(doc));

                // Several shape types can have an image including image shapes and OLE objects
                if (curNode.NodeType == NodeType.Shape && ((Shape)curNode).HasImage)
                    curNode.Remove();
                
                curNode = nextNode;
            }

            // The only remaining shape doesn't have an image
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, true).Count);
            Assert.False(((Shape)doc.GetChild(NodeType.Shape, 0, true)).HasImage);
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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default, the image is inserted at 100% scale
            Shape shape = builder.InsertImage(ImageDir + "Logo.jpg");

            // Reduce the overall size of the shape by 50%
            shape.Width = shape.Width * 0.5;
            shape.Height = shape.Height * 0.5;

            Assert.AreEqual(75.0d, shape.Width);
            Assert.AreEqual(75.0d, shape.Height);

            // However, we can also go back to the original image size and scale from there, for example, to 110%
            ImageSize imageSize = shape.ImageData.ImageSize;
            shape.Width = imageSize.WidthPoints * 1.1;
            shape.Height = imageSize.HeightPoints * 1.1;

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