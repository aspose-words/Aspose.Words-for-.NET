// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Settings;
using NUnit.Framework;
#if NET462 || JAVA
using System.Drawing;
using System.Drawing.Imaging;
#elif NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDocumentBuilderImages : ApiExampleBase
    {
        [Test]
        public void InsertImageFromStream()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream)
            //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert an image from a stream into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (Stream stream = File.OpenRead(ImageDir + "Logo.jpg"))
            {
                // Below are three ways of inserting an image from a stream.
                // 1 -  Inline shape with a default size based on the image's original dimensions:
                builder.InsertImage(stream);

                builder.InsertBreak(BreakType.PageBreak);

                // 2 -  Inline shape with custom dimensions:
                builder.InsertImage(stream, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

                builder.InsertBreak(BreakType.PageBreak);

                // 3 -  Floating shape with custom dimensions:
                builder.InsertImage(stream, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
                    100, 200, 100, WrapType.Square);
            }

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromStream.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilderImages.InsertImageFromStream.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(300.0d, imageShape.Height);
            Assert.AreEqual(300.0d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(108.0d, imageShape.Height);
            Assert.AreEqual(187.5d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.AreEqual(100.0d, imageShape.Height);
            Assert.AreEqual(200.0d, imageShape.Width);
            Assert.AreEqual(100.0d, imageShape.Left);
            Assert.AreEqual(100.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Square, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints);
        }

        [Test]
        public void InsertImageFromFilename()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
            //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert an image from the local file system into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are three ways of inserting an image from a local system filename.
            // 1 -  Inline shape with a default size based on the image's original dimensions:
            builder.InsertImage(ImageDir + "Logo.jpg");

            builder.InsertBreak(BreakType.PageBreak);

            // 2 -  Inline shape with custom dimensions:
            builder.InsertImage(ImageDir + "Transparent background logo.png", ConvertUtil.PixelToPoint(250),
                ConvertUtil.PixelToPoint(144));

            builder.InsertBreak(BreakType.PageBreak);

            // 3 -  Floating shape with custom dimensions:
            builder.InsertImage(ImageDir + "Windows MetaFile.wmf", RelativeHorizontalPosition.Margin, 100, 
                RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square);

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromFilename.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilderImages.InsertImageFromFilename.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(300.0d, imageShape.Height);
            Assert.AreEqual(300.0d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(108.0d, imageShape.Height);
            Assert.AreEqual(187.5d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.AreEqual(100.0d, imageShape.Height);
            Assert.AreEqual(200.0d, imageShape.Width);
            Assert.AreEqual(100.0d, imageShape.Left);
            Assert.AreEqual(100.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Square, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, imageShape);
            Assert.AreEqual(400.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(400.0d, imageShape.ImageData.ImageSize.WidthPoints);
        }

        [Test]
        public void InsertSvgImage()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExSummary:Shows how to determine which image will be inserted.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImageDir + "Scalable Vector Graphics.svg");

            // Aspose.Words insert SVG image to the document as PNG with svgBlip extension
            // that contains the original vector SVG image representation.
            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertSvgImage.SvgWithSvgBlip.docx");

            // Aspose.Words insert SVG image to the document as PNG, just like Microsoft Word does for old format.
            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertSvgImage.Svg.doc");

            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2003);

            // Aspose.Words insert SVG image to the document as EMF metafile to keep the image in vector representation.
            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertSvgImage.Emf.docx");
            //ExEnd
        }

#if NET462 || JAVA
        [Test]
        public void InsertImageFromImageObject()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image)
            //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert an image from an object into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image image = Image.FromFile(ImageDir + "Logo.jpg");

            // Below are three ways of inserting an image from an Image object instance.
            // 1 -  Inline shape with a default size based on the image's original dimensions:
            builder.InsertImage(image);

            builder.InsertBreak(BreakType.PageBreak);

            // 2 -  Inline shape with custom dimensions:
            builder.InsertImage(image, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

            builder.InsertBreak(BreakType.PageBreak);

            // 3 -  Floating shape with custom dimensions:
            builder.InsertImage(image, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
            100, 200, 100, WrapType.Square);

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromImageObject.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilderImages.InsertImageFromImageObject.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(300.0d, imageShape.Height);
            Assert.AreEqual(300.0d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(108.0d, imageShape.Height);
            Assert.AreEqual(187.5d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.AreEqual(100.0d, imageShape.Height);
            Assert.AreEqual(200.0d, imageShape.Width);
            Assert.AreEqual(100.0d, imageShape.Left);
            Assert.AreEqual(100.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Square, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints);
        }

        [Test]
        public void InsertImageFromByteArray()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[])
            //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert an image from a byte array into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image image = Image.FromFile(ImageDir + "Logo.jpg");

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Png);
                byte[] imageByteArray = ms.ToArray();

                // Below are three ways of inserting an image from a byte array.
                // 1 -  Inline shape with a default size based on the image's original dimensions:
                builder.InsertImage(imageByteArray);

                builder.InsertBreak(BreakType.PageBreak);

                // 2 -  Inline shape with custom dimensions:
                builder.InsertImage(imageByteArray, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

                builder.InsertBreak(BreakType.PageBreak);

                // 3 -  Floating shape with custom dimensions:
                builder.InsertImage(imageByteArray, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 
                100, 200, 100, WrapType.Square);
            }

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromByteArray.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilderImages.InsertImageFromByteArray.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(300.0d, imageShape.Height, 0.1d);
            Assert.AreEqual(300.0d, imageShape.Width, 0.1d);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(108.0d, imageShape.Height);
            Assert.AreEqual(187.5d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.AreEqual(100.0d, imageShape.Height);
            Assert.AreEqual(200.0d, imageShape.Width);
            Assert.AreEqual(100.0d, imageShape.Left);
            Assert.AreEqual(100.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Square, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);
        }
#elif NETCOREAPP2_1 || __MOBILE__
        [Test]
        public void InsertImageFromImageObjectNetStandard2()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert an image from an object into a document (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Decoding the image will convert it to the .png format.
            using (SKBitmap bitmap = SKBitmap.Decode(ImageDir + "Logo.jpg"))
            {
                // Below are three ways of inserting an image from an Image object instance.
                // 1 -  Inline shape with a default size based on the image's original dimensions:
                builder.InsertImage(bitmap);

                builder.InsertBreak(BreakType.PageBreak);

                // 2 -  Inline shape with custom dimensions:
                builder.InsertImage(bitmap, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

                builder.InsertBreak(BreakType.PageBreak);

                // 3 -  Floating shape with custom dimensions:
                builder.InsertImage(bitmap, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
                    100, 200, 100, WrapType.Square);
            }

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromImageObjectNetStandard2.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilderImages.InsertImageFromImageObjectNetStandard2.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(300.0d, imageShape.Height, 0.1d);
            Assert.AreEqual(300.0d, imageShape.Width, 0.1d);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(108.0d, imageShape.Height);
            Assert.AreEqual(187.5d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.AreEqual(100.0d, imageShape.Height);
            Assert.AreEqual(200.0d, imageShape.Width);
            Assert.AreEqual(100.0d, imageShape.Left);
            Assert.AreEqual(100.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Square, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);
        }

        [Test]
        public void InsertImageFromByteArrayNetStandard2()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[])
            //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert an image from a byte array into a document (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Decoding the image will convert it to the .png format.
            using (SKBitmap bitmap = SKBitmap.Decode(ImageDir + "Logo.jpg"))
            {
                using (SKImage image = SKImage.FromBitmap(bitmap))
                {
                    using (SKData data = image.Encode())
                    {
                        byte[] imageByteArray = data.ToArray();

                        // Below are three ways of inserting an image from a byte array.
                        // 1 -  Inline shape with a default size based on the image's original dimensions:
                        builder.InsertImage(imageByteArray);

                        builder.InsertBreak(BreakType.PageBreak);

                        // 2 -  Inline shape with custom dimensions:
                        builder.InsertImage(imageByteArray, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

                        builder.InsertBreak(BreakType.PageBreak);

                        // 3 -  Floating shape with custom dimensions:
                        builder.InsertImage(imageByteArray, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
                            100, 200, 100, WrapType.Square);
                    }
                }
            }
            
            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromByteArrayNetStandard2.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilderImages.InsertImageFromByteArrayNetStandard2.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(300.0d, imageShape.Height, 0.1d);
            Assert.AreEqual(300.0d, imageShape.Width, 0.1d);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(108.0d, imageShape.Height);
            Assert.AreEqual(187.5d, imageShape.Width);
            Assert.AreEqual(0.0d, imageShape.Left);
            Assert.AreEqual(0.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Inline, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Column, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.AreEqual(100.0d, imageShape.Height);
            Assert.AreEqual(200.0d, imageShape.Width);
            Assert.AreEqual(100.0d, imageShape.Left);
            Assert.AreEqual(100.0d, imageShape.Top);

            Assert.AreEqual(WrapType.Square, imageShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, imageShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, imageShape.RelativeVerticalPosition);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.HeightPoints, 0.1d);
            Assert.AreEqual(300.0d, imageShape.ImageData.ImageSize.WidthPoints, 0.1d);
        }
#endif
    }
}