// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Document
{
    [TestFixture]
    public class ExDocumentBuilderImages : QaTestsBase
    {
        [Test]
        public void InsertImageStreamRelativePositionEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to import an image from a stream into a document using relative positions.
            DocumentBuilder builder = new DocumentBuilder();

            System.IO.Stream stream = System.IO.File.OpenRead(ExDir + "Aspose.Words.gif");
            try
            {
                builder.InsertImage(stream, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100,
                                    200, 100, WrapType.Square);
            }
            finally
            { 
                stream.Close();
            }


            builder.Document.Save(ExDir + "Image.CreateFromStreamRelativePosition Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromByteArrayEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[])
            //ExSummary:Shows how to import an image from a byte array into a document.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder();

            // Prepare a byte array of an image.
            System.Drawing.Image image = System.Drawing.Image.FromFile(ExDir + "Aspose.Words.gif");
            System.Drawing.ImageConverter imageConverter = new System.Drawing.ImageConverter();
            byte[] imageBytes = (byte[])imageConverter.ConvertTo(image, typeof (byte[]));

            builder.InsertImage(imageBytes);
            builder.Document.Save(ExDir + "Image.CreateFromByteArray Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromByteArrayDoubleEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
            //ExSummary:Shows how to import an image from a byte array into a document with a custom size.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder();

            // Prepare a byte array of an image.
            System.Drawing.Image image = System.Drawing.Image.FromFile(ExDir + "Aspose.Words.gif");
            System.Drawing.ImageConverter imageConverter = new System.Drawing.ImageConverter();
            byte[] imageBytes = (byte[])imageConverter.ConvertTo(image, typeof(byte[]));

            builder.InsertImage(imageBytes, Aspose.Words.ConvertUtil.PixelToPoint(450), Aspose.Words.ConvertUtil.PixelToPoint(144));
            builder.Document.Save(ExDir + "Image.CreateFromByteArrayDouble Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromByteArrayRelativePositionEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to import an image from a byte array into a document using relative positions.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder();

            // Prepare a byte array of an image.
            System.Drawing.Image image = System.Drawing.Image.FromFile(ExDir + "Aspose.Words.gif");
            System.Drawing.ImageConverter imageConverter = new System.Drawing.ImageConverter();
            byte[] imageBytes = (byte[])imageConverter.ConvertTo(image, typeof(byte[]));

            builder.InsertImage(imageBytes, Aspose.Words.ConvertUtil.PixelToPoint(450), Aspose.Words.ConvertUtil.PixelToPoint(144));
            builder.Document.Save(ExDir + "Image.CreateFromByteArrayRelativePosition Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromImageDoubleEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
            //ExSummary:Shows how to import an image into a document with a custom size.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder();

            System.Drawing.Image rasterImage = System.Drawing.Image.FromFile(ExDir + "Aspose.Words.gif");
            try
            {
                builder.InsertImage(rasterImage,
                                    Aspose.Words.ConvertUtil.PixelToPoint(450), Aspose.Words.ConvertUtil.PixelToPoint(144));
                builder.Writeln();
            }
            finally
            {
                rasterImage.Dispose();
            }
            builder.Document.Save(ExDir + "Image.CreateFromImageWithStreamCustomSize Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromImageRelativePositionEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to import an image from a stream into a document using relative positions.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder();

            System.Drawing.Image rasterImage = System.Drawing.Image.FromFile(ExDir + "Aspose.Words.gif");
            try
            {
                builder.InsertImage(rasterImage, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100,
                                    200, 100, WrapType.Square);
            }
            finally
            {
                rasterImage.Dispose();
            }

            builder.Document.Save(ExDir + "Image.CreateFromImageWithStreamRelativePosition Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageStreamDoubleEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
            //ExSummary:Shows how to import an image from a stream into a document with a custom size.
            DocumentBuilder builder = new DocumentBuilder();

            System.IO.Stream stream = System.IO.File.OpenRead(ExDir + "Aspose.Words.gif");
            try
            {
                builder.InsertImage(stream, Aspose.Words.ConvertUtil.PixelToPoint(400), Aspose.Words.ConvertUtil.PixelToPoint(400));
            }
            finally
            {
                stream.Close();
            }

            builder.Document.Save(ExDir + "Image.CreateFromStreamCustomSize Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageStringDoubleDoubleEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
            //ExSummary:Shows how to import an image from a url into a document with a custom size.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Remote URI
            builder.InsertImage("http://www.aspose.com/images/aspose-logo.gif",
                Aspose.Words.ConvertUtil.PixelToPoint(450), Aspose.Words.ConvertUtil.PixelToPoint(144));

            // Local URI
            builder.InsertImage(ExDir + "Aspose.Words.gif",
                Aspose.Words.ConvertUtil.PixelToPoint(400), Aspose.Words.ConvertUtil.PixelToPoint(400));

            doc.Save(ExDir + "DocumentBuilder.InsertImageFromUrlCustomSize Out.doc");
            //ExEnd
        }
    }
}
