//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace Examples
{
    /// <summary>
    /// Mostly scenarios that deal with image shapes.
    /// </summary>
    [TestFixture]
    public class ExImage : ExBase
    {
        [Test]
        public void CreateFromUrl()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(string)
            //ExFor:DocumentBuilder.Writeln
            //ExSummary:Shows how to inserts an image from a URL. The image is inserted inline and at 100% scale.
            // This creates a builder and also an empty document inside the builder.
            DocumentBuilder builder = new DocumentBuilder();

            builder.Write("Image from local file: ");
            builder.InsertImage(MyDir + "Aspose.Words.gif");
            builder.Writeln();

            builder.Write("Image from an internet url, automatically downloaded for you: ");
            builder.InsertImage("http://www.aspose.com/Images/aspose-logo.jpg");
            builder.Writeln();

            builder.Document.Save(MyDir + "Image.CreateFromUrl Out.doc");
            //ExEnd
        }

        [Test]
        public void CreateFromStream()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream)
            //ExSummary:Shows how to insert an image from a stream. The image is inserted inline and at 100% scale.
            // This creates a builder and also an empty document inside the builder.
            DocumentBuilder builder = new DocumentBuilder();

            Stream stream = File.OpenRead(MyDir + "Aspose.Words.gif");
            try
            {
                builder.Write("Image from stream: ");
                builder.InsertImage(stream);
            }
            finally
            {
                stream.Close();
            }

            builder.Document.Save(MyDir + "Image.CreateFromStream Out.doc");
            //ExEnd
        }

        [Test]
        public void CreateFromImage()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image)
            //ExSummary:Shows how to insert a .NET Image object into a document. The image is inserted inline and at 100% scale.
            // This creates a builder and also an empty document inside the builder.
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a raster image.
            System.Drawing.Image rasterImage = System.Drawing.Image.FromFile(MyDir + "Aspose.Words.gif");
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

            // Aspose.Words allows to insert a metafile too.
            System.Drawing.Image metafile = System.Drawing.Image.FromFile(MyDir + "Hammer.wmf");
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

            builder.Document.Save(MyDir + "Image.CreateFromImage Out.doc");
            //ExEnd
        }

        [Test]
        public void CreateFloatingPageCenter()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(string)
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
            // This creates a builder and also an empty document inside the builder.
            DocumentBuilder builder = new DocumentBuilder();

            // By default, the image is inline.
            Shape shape = builder.InsertImage(MyDir + "Aspose.Words.gif");

            // Make the image float, put it behind text and center on the page.
            shape.WrapType = WrapType.None;
            shape.BehindText = true;
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.HorizontalAlignment = HorizontalAlignment.Center;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            shape.VerticalAlignment = VerticalAlignment.Center;

            builder.Document.Save(MyDir + "Image.CreateFloatingPageCenter Out.doc");
            //ExEnd
        }

        [Test]
        public void CreateFloatingPositionSize()
        {
            //ExStart
            //ExFor:ShapeBase.Left
            //ExFor:ShapeBase.Top
            //ExFor:ShapeBase.Width
            //ExFor:ShapeBase.Height
            //ExFor:DocumentBuilder.CurrentSection
            //ExFor:PageSetup.PageWidth
            //ExSummary:Shows how to insert a floating image and specify its position and size.
            // This creates a builder and also an empty document inside the builder.
            DocumentBuilder builder = new DocumentBuilder();

            // By default, the image is inline.
            Shape shape = builder.InsertImage(MyDir + "Hammer.wmf");

            // Make the image float, put it behind text and center on the page.
            shape.WrapType = WrapType.None;

            // Make position relative to the page.
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Make the shape occupy a band 50 points high at the very top of the page.
            shape.Left = 0;
            shape.Top = 0;
            shape.Width = builder.CurrentSection.PageSetup.PageWidth;
            shape.Height = 50;

            builder.Document.Save(MyDir + "Image.CreateFloatingPositionSize Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageWithHyperlink()
        {
            //ExStart
            //ExFor:ShapeBase.HRef
            //ExFor:ShapeBase.ScreenTip
            //ExSummary:Shows how to insert an image with a hyperlink.
            // This creates a builder and also an empty document inside the builder.
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertImage(MyDir + "Hammer.wmf");
            shape.HRef = "http://www.aspose.com/Community/Forums/75/ShowForum.aspx";
            shape.ScreenTip = "Aspose.Words Support Forums";

            builder.Document.Save(MyDir + "Image.InsertImageWithHyperlink Out.doc");
            //ExEnd
        }

        [Test]
        public void CreateImageDirectly()
        {
            //ExStart
            //ExFor:Shape.#ctor(DocumentBase,ShapeType)
            //ExFor:ShapeType
            //ExSummary:Shows how to create and add an image to a document without using document builder.
            Document doc = new Document();

            Shape shape = new Shape(doc, ShapeType.Image);
            shape.ImageData.SetImage(MyDir + "Hammer.wmf");
            shape.Width = 100;
            shape.Height = 100;

            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            doc.Save(MyDir + "Image.CreateImageDirectly Out.doc");
            //ExEnd
        }

        [Test]
        public void CreateLinkedImage()
        {
            //ExStart
            //ExFor:Shape.ImageData
            //ExFor:ImageData
            //ExFor:ImageData.SourceFullName
            //ExFor:ImageData.SetImage(string)
            //ExFor:DocumentBuilder.InsertNode
            //ExSummary:Shows how to insert a linked image into a document. 
            DocumentBuilder builder = new DocumentBuilder();

            string imageFileName = MyDir + "Hammer.wmf";

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

            builder.Document.Save(MyDir + "Image.CreateLinkedImage Out.doc");
            //ExEnd
        }

        [Test]
        public void DeleteAllImages()
        {
            Document doc = new Document(MyDir + "Image.SampleImages.doc");
            Assert.AreEqual(6, doc.GetChildNodes(NodeType.Shape, true).Count);
            
            //ExStart
            //ExFor:Shape.HasImage
            //ExFor:Node.Remove
            //ExSummary:Shows how to delete all images from a document.
            // Here we get all shapes from the document node, but you can do this for any smaller
            // node too, for example delete shapes from a single section or a paragraph.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            // We cannot delete shape nodes while we enumerate through the collection.
            // One solution is to add nodes that we want to delete to a temporary array and delete afterwards.
            ArrayList shapesToDelete = new ArrayList();
            foreach (Shape shape in shapes)
            {
                // Several shape types can have an image including image shapes and OLE objects.
                if (shape.HasImage)
                    shapesToDelete.Add(shape);
            }

            // Now we can delete shapes.
            foreach (Shape shape in shapesToDelete)
                shape.Remove();
            //ExEnd

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, true).Count);
            doc.Save(MyDir + "Image.DeleteAllImages Out.doc");
        }

        [Test]
        public void DeleteAllImagesPreOrder()
        {
            Document doc = new Document(MyDir + "Image.SampleImages.doc");
            Assert.AreEqual(6, doc.GetChildNodes(NodeType.Shape, true).Count);
            
            //ExStart
            //ExFor:Node.NextPreOrder
            //ExSummary:Shows how to delete all images from a document using pre-order tree traversal.
            Node curNode = doc;
            while (curNode != null)
            {
                Node nextNode = curNode.NextPreOrder(doc);

                if (curNode.NodeType.Equals(NodeType.Shape))
                {
                    Shape shape = (Shape)curNode;

                    // Several shape types can have an image including image shapes and OLE objects.
                    if (shape.HasImage)
                        shape.Remove();
                }

                curNode = nextNode;
            }
            //ExEnd

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Shape, true).Count);
            doc.Save(MyDir + "Image.DeleteAllImagesPreOrder Out.doc");
        }

        //ExStart
        //ExFor:Shape
        //ExFor:Shape.ImageData
        //ExFor:Shape.HasImage
        //ExFor:ImageData
        //ExFor:FileFormatUtil.ImageTypeToExtension(Aspose.Words.Drawing.ImageType)
        //ExFor:ImageData.ImageType
        //ExFor:ImageData.Save(string)
        //ExFor:DrawingMLImageData
        //ExFor:DrawingMLImageData.ImageType
        //ExFor:DrawingMLImageData.Save(string)
        //ExFor:CompositeNode.GetChildNodes(NodeType, bool)
        //ExId:ExtractImagesToFiles
        //ExSummary:Shows how to extract images from a document and save them as files.
        [Test] //ExSkip
        public void ExtractImagesToFiles()
        {
            Document doc = new Document(MyDir + "Image.SampleImages.doc");

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;			
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = string.Format(
                        "Image.ExportImages.{0} Out{1}", imageIndex, FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType));
                    shape.ImageData.Save(MyDir + imageFileName);
                    imageIndex++;
                }
            }

            // Newer Microsoft Word documents (such as DOCX) may contain a different type of image container called DrawingML.
            // Repeat the process to extract these if they are present in the loaded document.
            NodeCollection dmlShapes = doc.GetChildNodes(NodeType.DrawingML, true);
            foreach (DrawingML dml in dmlShapes)
            {
                if (dml.HasImage)
                {
                    string imageFileName = string.Format(
                        "Image.ExportImages.{0} Out{1}", imageIndex, FileFormatUtil.ImageTypeToExtension(dml.ImageData.ImageType));
                    dml.ImageData.Save(MyDir + imageFileName);
                    imageIndex++;
                }
            }
        }
        //ExEnd

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

            // By default, the image is inserted at 100% scale.
            Shape shape = builder.InsertImage(MyDir + "Aspose.Words.gif");

            // It is easy to change the shape size. In this case, make it 50% relative to the current shape size.
            shape.Width = shape.Width * 0.5;
            shape.Height = shape.Height * 0.5;

            // However, we can also go back to the original image size and scale from there, say 110%.
            ImageSize imageSize = shape.ImageData.ImageSize;
            shape.Width = imageSize.WidthPoints * 1.1;
            shape.Height = imageSize.HeightPoints * 1.1;

            builder.Document.Save(MyDir + "Image.ScaleImage Out.doc");
            //ExEnd
        }
    }
}
