// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;
#if NETFRAMEWORK
using System.Drawing;
#else
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
        public void CreateFromUrl()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExSummary:Shows how to inserts an image from a URL. The image is inserted inline and at 100% scale.
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            builder.Write("Image from local file: ");
            builder.InsertImage(ImageDir + "Aspose.Words.gif");
            builder.Writeln();

            builder.Write("Image from an Internet url, automatically downloaded for you: ");
            builder.InsertImage(AsposeLogoUrl);
            builder.Writeln();

            builder.Document.Save(ArtifactsDir + "Image.CreateFromUrl.doc");
            //ExEnd
        }

        [Test]
        public void CreateFromStream()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream)
            //ExSummary:Shows how to insert an image from a stream. The image is inserted inline and at 100% scale.
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            Stream stream = File.OpenRead(ImageDir + "Aspose.Words.gif");
            try
            {
                builder.Write("Image from stream: ");
                builder.InsertImage(stream);
            }
            finally
            {
                stream.Close();
            }

            builder.Document.Save(ArtifactsDir + "Image.CreateFromStream.doc");
            //ExEnd
        }

        #if NETFRAMEWORK
        [Test]
        [Category("SkipMono")]
        public void CreateFromImage()
        {
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a raster image
            Image rasterImage = Image.FromFile(ImageDir + "Aspose.Words.gif");
            try
            {
                builder.Write("Raster image: ");
                builder.InsertImage(rasterImage);
                builder.Writeln();
            }
            finally
            {
                rasterImage.Dispose();
            }

            // Aspose.Words allows to insert a metafile too
            Image metafile = Image.FromFile(ImageDir + "Hammer.wmf");
            try
            {
                builder.Write("Metafile: ");
                builder.InsertImage(metafile);
                builder.Writeln();
            }
            finally
            {
                metafile.Dispose();
            }

            builder.Document.Save(ArtifactsDir + "Image.CreateFromImage.doc");
        }
        #else
        [Test]
        [Category("SkipMono")]
        public void CreateFromImageNetStandard2()
        {
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a raster image
            // SKBitmap doesn't allow to insert a metafiles
            using (SKBitmap rasterImage = SKBitmap.Decode(ImageDir + "Aspose.Words.gif"))
            {
                builder.Write("Raster image: ");
                builder.InsertImage(rasterImage);
                builder.Writeln();
            }

            builder.Document.Save(ArtifactsDir + "Image.CreateFromImage.doc");
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
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            // By default, the image is inline
            Shape shape = builder.InsertImage(ImageDir + "Aspose.Words.gif");

            // Make the image float, put it behind text and center on the page
            shape.WrapType = WrapType.None;
            shape.BehindText = true;
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.HorizontalAlignment = HorizontalAlignment.Center;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            shape.VerticalAlignment = VerticalAlignment.Center;

            builder.Document.Save(ArtifactsDir + "Image.CreateFloatingPageCenter.doc");
            //ExEnd
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
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            // By default, the image is inline
            Shape shape = builder.InsertImage(ImageDir + "Aspose.Words.gif");

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

            builder.Document.Save(ArtifactsDir + "Image.CreateFloatingPositionSize.docx");
            //ExEnd
        }

        [Test]
        public void InsertImageWithHyperlink()
        {
            //ExStart
            //ExFor:ShapeBase.HRef
            //ExFor:ShapeBase.ScreenTip
            //ExFor:ShapeBase.Target
            //ExSummary:Shows how to insert an image with a hyperlink.
            // This creates a builder and also an empty document inside the builder
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertImage(ImageDir + "Hammer.wmf");
            shape.HRef = "http://www.aspose.com/Community/Forums/75/ShowForum.aspx";
            shape.Target = "New Window";
            shape.ScreenTip = "Aspose.Words Support Forums";

            builder.Document.Save(ArtifactsDir + "Image.InsertImageWithHyperlink.doc");
            //ExEnd
        }

        [Test]
        public void CreateImageDirectly()
        {
            //ExStart
            //ExFor:Shape.#ctor(DocumentBase,ShapeType)
            //ExFor:ShapeType
            //ExSummary:Shows how to create shape and add an image to a document without using a document builder.
            Document doc = new Document();

            // Public constructor of "Shape" class creates shape with "ShapeMarkupLanguage.Vml" markup type
            // If you need to create "NonPrimitive" shapes, like SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
            // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
            // please use DocumentBuilder.InsertShape methods
            Shape shape = new Shape(doc, ShapeType.Image);
            shape.ImageData.SetImage(ImageDir + "Hammer.wmf");
            shape.Width = 100;
            shape.Height = 100;

            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            doc.Save(ArtifactsDir + "Image.CreateImageDirectly.doc");
            //ExEnd
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
            DocumentBuilder builder = new DocumentBuilder();

            string imageFileName = ImageDir + "Hammer.wmf";

            builder.Write("Image linked, not stored in the document: ");

            Shape linkedOnly = new Shape(builder.Document, ShapeType.Image);
            linkedOnly.WrapType = WrapType.Inline;
            linkedOnly.ImageData.SourceFullName = imageFileName;

            builder.InsertNode(linkedOnly);
            builder.Writeln();

            builder.Write("Image linked and stored in the document: ");

            Shape linkedAndStored = new Shape(builder.Document, ShapeType.Image);
            linkedAndStored.WrapType = WrapType.Inline;
            linkedAndStored.ImageData.SourceFullName = imageFileName;
            linkedAndStored.ImageData.SetImage(imageFileName);

            builder.InsertNode(linkedAndStored);
            builder.Writeln();

            builder.Write("Image stored in the document, but not linked: ");

            Shape stored = new Shape(builder.Document, ShapeType.Image);
            stored.WrapType = WrapType.Inline;
            stored.ImageData.SetImage(imageFileName);

            builder.InsertNode(stored);
            builder.Writeln();

            builder.Document.Save(ArtifactsDir + "Image.CreateLinkedImage.doc");
            //ExEnd
        }

        [Test]
        public void DeleteAllImages()
        {
            //ExStart
            //ExFor:Shape.HasImage
            //ExFor:Node.Remove
            //ExSummary:Shows how to delete all images from a document.
            Document doc = new Document(MyDir + "SampleImages.docx");
            Assert.AreEqual(6, doc.GetChildNodes(NodeType.Shape, true).Count);

            // Here we get all shapes from the document node, but you can do this for any smaller
            // node too, for example delete shapes from a single section or a paragraph
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            // We cannot delete shape nodes while we enumerate through the collection
            // One solution is to add nodes that we want to delete to a temporary array and delete afterwards
            ArrayList shapesToDelete = new ArrayList();
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                // Several shape types can have an image including image shapes and OLE objects
                if (shape.HasImage)
                    shapesToDelete.Add(shape);
            }

            // Now we can delete shapes
            foreach (Shape shape in shapesToDelete)
                shape.Remove();

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, true).Count);
            doc.Save(ArtifactsDir + "Image.DeleteAllImages.docx");
            //ExEnd
        }

        [Test]
        public void DeleteAllImagesPreOrder()
        {
            //ExStart
            //ExFor:Node.NextPreOrder(Node)
            //ExFor:Node.PreviousPreOrder(Node)
            //ExSummary:Shows how to delete all images from a document using pre-order tree traversal.
            Document doc = new Document(MyDir + "SampleImages.docx");
            Assert.AreEqual(6, doc.GetChildNodes(NodeType.Shape, true).Count);

            Node curNode = doc;
            while (curNode != null)
            {
                Node nextNode = curNode.NextPreOrder(doc);

                if (curNode.PreviousPreOrder(doc) != null && nextNode != null)
                {
                    Assert.AreEqual(curNode, nextNode.PreviousPreOrder(doc));
                }

                if (curNode.NodeType.Equals(NodeType.Shape))
                {
                    Shape shape = (Shape) curNode;

                    // Several shape types can have an image including image shapes and OLE objects
                    if (shape.HasImage)
                        shape.Remove();
                }

                curNode = nextNode;
            }

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, true).Count);
            doc.Save(ArtifactsDir + "Image.DeleteAllImagesPreOrder.docx");
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
            //ExSummary:Shows how to resize an image shape.
            DocumentBuilder builder = new DocumentBuilder();

            // By default, the image is inserted at 100% scale
            Shape shape = builder.InsertImage(ImageDir + "Aspose.Words.gif");

            // It is easy to change the shape size. In this case, make it 50% relative to the current shape size
            shape.Width = shape.Width * 0.5;
            shape.Height = shape.Height * 0.5;

            // However, we can also go back to the original image size and scale from there, say 110%
            ImageSize imageSize = shape.ImageData.ImageSize;
            shape.Width = imageSize.WidthPoints * 1.1;
            shape.Height = imageSize.HeightPoints * 1.1;

            builder.Document.Save(ArtifactsDir + "Image.ScaleImage.doc");
            //ExEnd
        }
    }
}