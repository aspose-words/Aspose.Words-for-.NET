﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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

            Assert.That(imageShape.Height, Is.EqualTo(300.0d));
            Assert.That(imageShape.Width, Is.EqualTo(300.0d));
            Assert.That(imageShape.Left, Is.EqualTo(0.0d));
            Assert.That(imageShape.Top, Is.EqualTo(0.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Column));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d));

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.That(imageShape.Height, Is.EqualTo(108.0d));
            Assert.That(imageShape.Width, Is.EqualTo(187.5d));
            Assert.That(imageShape.Left, Is.EqualTo(0.0d));
            Assert.That(imageShape.Top, Is.EqualTo(0.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Column));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d));

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.That(imageShape.Height, Is.EqualTo(100.0d));
            Assert.That(imageShape.Width, Is.EqualTo(200.0d));
            Assert.That(imageShape.Left, Is.EqualTo(100.0d));
            Assert.That(imageShape.Top, Is.EqualTo(100.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Square));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Margin));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Margin));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d));
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

            Assert.That(imageShape.Height, Is.EqualTo(300.0d));
            Assert.That(imageShape.Width, Is.EqualTo(300.0d));
            Assert.That(imageShape.Left, Is.EqualTo(0.0d));
            Assert.That(imageShape.Top, Is.EqualTo(0.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Column));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d));

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.That(imageShape.Height, Is.EqualTo(108.0d));
            Assert.That(imageShape.Width, Is.EqualTo(187.5d));
            Assert.That(imageShape.Left, Is.EqualTo(0.0d));
            Assert.That(imageShape.Top, Is.EqualTo(0.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Column));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d));

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.That(imageShape.Height, Is.EqualTo(100.0d));
            Assert.That(imageShape.Width, Is.EqualTo(200.0d));
            Assert.That(imageShape.Left, Is.EqualTo(100.0d));
            Assert.That(imageShape.Top, Is.EqualTo(100.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Square));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Margin));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Margin));

            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(400.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(400.0d));
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

            string imageFile = ImageDir + "Logo.jpg";

            // Below are three ways of inserting an image from an Image object instance.
            // 1 -  Inline shape with a default size based on the image's original dimensions:
            builder.InsertImage(imageFile);

            builder.InsertBreak(BreakType.PageBreak);

            // 2 -  Inline shape with custom dimensions:
            builder.InsertImage(imageFile, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

            builder.InsertBreak(BreakType.PageBreak);

            // 3 -  Floating shape with custom dimensions:
            builder.InsertImage(imageFile, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
            100, 200, 100, WrapType.Square);

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromImageObject.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilderImages.InsertImageFromImageObject.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(imageShape.Height, Is.EqualTo(300.0d));
            Assert.That(imageShape.Width, Is.EqualTo(300.0d));
            Assert.That(imageShape.Left, Is.EqualTo(0.0d));
            Assert.That(imageShape.Top, Is.EqualTo(0.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Column));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d));

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.That(imageShape.Height, Is.EqualTo(108.0d));
            Assert.That(imageShape.Width, Is.EqualTo(187.5d));
            Assert.That(imageShape.Left, Is.EqualTo(0.0d));
            Assert.That(imageShape.Top, Is.EqualTo(0.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Column));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d));

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.That(imageShape.Height, Is.EqualTo(100.0d));
            Assert.That(imageShape.Width, Is.EqualTo(200.0d));
            Assert.That(imageShape.Left, Is.EqualTo(100.0d));
            Assert.That(imageShape.Top, Is.EqualTo(100.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Square));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Margin));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Margin));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d));
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

            byte[] imageByteArray = TestUtil.ImageToByteArray(ImageDir + "Logo.jpg");

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

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromByteArray.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilderImages.InsertImageFromByteArray.docx");

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(imageShape.Height, Is.EqualTo(300.0d).Within(0.1d));
            Assert.That(imageShape.Width, Is.EqualTo(300.0d).Within(0.1d));
            Assert.That(imageShape.Left, Is.EqualTo(0.0d));
            Assert.That(imageShape.Top, Is.EqualTo(0.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Column));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d).Within(0.1d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d).Within(0.1d));

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.That(imageShape.Height, Is.EqualTo(108.0d));
            Assert.That(imageShape.Width, Is.EqualTo(187.5d));
            Assert.That(imageShape.Left, Is.EqualTo(0.0d));
            Assert.That(imageShape.Top, Is.EqualTo(0.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Inline));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Column));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Paragraph));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d).Within(0.1d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d).Within(0.1d));

            imageShape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.That(imageShape.Height, Is.EqualTo(100.0d));
            Assert.That(imageShape.Width, Is.EqualTo(200.0d));
            Assert.That(imageShape.Left, Is.EqualTo(100.0d));
            Assert.That(imageShape.Top, Is.EqualTo(100.0d));

            Assert.That(imageShape.WrapType, Is.EqualTo(WrapType.Square));
            Assert.That(imageShape.RelativeHorizontalPosition, Is.EqualTo(RelativeHorizontalPosition.Margin));
            Assert.That(imageShape.RelativeVerticalPosition, Is.EqualTo(RelativeVerticalPosition.Margin));

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            Assert.That(imageShape.ImageData.ImageSize.HeightPoints, Is.EqualTo(300.0d).Within(0.1d));
            Assert.That(imageShape.ImageData.ImageSize.WidthPoints, Is.EqualTo(300.0d).Within(0.1d));
        }

        [Test]
        public void InsertGif()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExSummary:Shows how to insert gif image to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // We can insert gif image using path or bytes array.
            // It works only if DocumentBuilder optimized to Word version 2010 or higher.
            // Note, that access to the image bytes causes conversion Gif to Png.
            Shape gifImage = builder.InsertImage(ImageDir + "Graphics Interchange Format.gif");

            gifImage = builder.InsertImage(File.ReadAllBytes(ImageDir + "Graphics Interchange Format.gif"));

            builder.Document.Save(ArtifactsDir + "InsertGif.docx");
            //ExEnd
        }
    }
}