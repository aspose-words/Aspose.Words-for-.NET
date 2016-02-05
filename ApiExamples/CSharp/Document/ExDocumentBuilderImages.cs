// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;


namespace ApiExamples.Document
{
    [TestFixture]
    public class ExDocumentBuilderImages : ApiExampleBase
    {
        [Test]
        public void InsertImageStreamRelativePositionEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert an image into a document from a stream, also using relative positions.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            System.IO.Stream stream = System.IO.File.OpenRead(MyDir + "Aspose.Words.gif");
            try
            {
                builder.InsertImage(stream, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100,
                                    200, 100, WrapType.Square);
            }
            finally
            { 
                stream.Close();
            }

            builder.Document.Save(MyDir + "Image.CreateFromStreamRelativePosition Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromByteArrayEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[])
            //ExSummary:Shows how to import an image into a document from a byte array.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare a byte array of an image.
            System.Drawing.Image image = System.Drawing.Image.FromFile(MyDir + "Aspose.Words.gif");
            System.Drawing.ImageConverter imageConverter = new System.Drawing.ImageConverter();
            byte[] imageBytes = (byte[])imageConverter.ConvertTo(image, typeof (byte[]));

            builder.InsertImage(imageBytes);
            builder.Document.Save(MyDir + "Image.CreateFromByteArrayDefault Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromByteArrayCustomSizeEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
            //ExSummary:Shows how to import an image into a document from a byte array, with a custom size.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare a byte array of an image.
            System.Drawing.Image image = System.Drawing.Image.FromFile(MyDir + "Aspose.Words.gif");
            System.Drawing.ImageConverter imageConverter = new System.Drawing.ImageConverter();
            byte[] imageBytes = (byte[])imageConverter.ConvertTo(image, typeof(byte[]));

            builder.InsertImage(imageBytes, Aspose.Words.ConvertUtil.PixelToPoint(450), Aspose.Words.ConvertUtil.PixelToPoint(144));
            builder.Document.Save(MyDir + "Image.CreateFromByteArrayCustomSize Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromByteArrayRelativePositionEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to import an image into a document from a byte array, also using relative positions.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare a byte array of an image.
            System.Drawing.Image image = System.Drawing.Image.FromFile(MyDir + "Aspose.Words.gif");
            System.Drawing.ImageConverter imageConverter = new System.Drawing.ImageConverter();
            byte[] imageBytes = (byte[])imageConverter.ConvertTo(image, typeof(byte[]));

            builder.InsertImage(imageBytes, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100,
                                    200, 100, WrapType.Square);
            builder.Document.Save(MyDir + "Image.CreateFromByteArrayRelativePosition Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromImageCustomSizeEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
            //ExSummary:Shows how to import an image into a document, with a custom size.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            System.Drawing.Image rasterImage = System.Drawing.Image.FromFile(MyDir + "Aspose.Words.gif");
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
            builder.Document.Save(MyDir + "Image.CreateFromImageWithStreamCustomSize Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageFromImageRelativePositionEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to import an image into a document, also using relative positions.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            System.Drawing.Image rasterImage = System.Drawing.Image.FromFile(MyDir + "Aspose.Words.gif");
            try
            {
                builder.InsertImage(rasterImage, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100,
                                    200, 100, WrapType.Square);
            }
            finally
            {
                rasterImage.Dispose();
            }

            builder.Document.Save(MyDir + "Image.CreateFromImageWithStreamRelativePosition Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageStreamCustomSizeEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
            //ExSummary:Shows how to import an image from a stream into a document with a custom size.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            System.IO.Stream stream = System.IO.File.OpenRead(MyDir + "Aspose.Words.gif");
            try
            {
                builder.InsertImage(stream, Aspose.Words.ConvertUtil.PixelToPoint(400), Aspose.Words.ConvertUtil.PixelToPoint(400));
            }
            finally
            {
                stream.Close();
            }

            builder.Document.Save(MyDir + "Image.CreateFromStreamCustomSize Out.doc");
            //ExEnd
        }

        [Test]
        public void InsertImageStringCustomSizeEx()
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
            builder.InsertImage(MyDir + "Aspose.Words.gif",
                Aspose.Words.ConvertUtil.PixelToPoint(400), Aspose.Words.ConvertUtil.PixelToPoint(400));

            doc.Save(MyDir + "DocumentBuilder.InsertImageFromUrlCustomSize Out.doc");
            //ExEnd
        }
    }
}
