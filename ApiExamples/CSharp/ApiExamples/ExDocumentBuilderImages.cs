// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;
#if NETFRAMEWORK
using System.Drawing;
using System.Drawing.Imaging;
#else
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
            //ExSummary:Shows different solutions of how to import an image into a document from a stream.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (Stream stream = File.OpenRead(ImageDir + "Logo.jpg"))
            {
                builder.Writeln("Inserted image from stream: ");
                builder.InsertImage(stream);
                
                builder.Writeln("\nInserted image from stream with a custom size: ");
                builder.InsertImage(stream, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));
                
                builder.Writeln("\nInserted image from stream using relative positions: ");
                builder.InsertImage(stream, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
                    100, 200, 100, WrapType.Square);
            }

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromStream.docx");
            //ExEnd
        }

        [Test]
        public void InsertImageFromString()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
            //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows different solutions of how to import an image into a document from a string.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("\nInserted image from string: ");
            builder.InsertImage(ImageDir + "Logo.jpg");

            builder.Writeln("\nInserted image from string with a custom size: ");
            builder.InsertImage(ImageDir + "Logo.jpg", ConvertUtil.PixelToPoint(250),
                ConvertUtil.PixelToPoint(144));

            builder.Writeln("\nInserted image from string using relative positions: ");
            builder.InsertImage(ImageDir + "Logo.jpg", RelativeHorizontalPosition.Margin, 100, 
                RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square);

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromString.docx");
            //ExEnd
        }

        #if NETFRAMEWORK
        [Test]
        public void InsertImageFromImageClass()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows different solutions of how to import an image into a document from Image class.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image image = Image.FromFile(ImageDir + "Logo.jpg");

            builder.Writeln("\nInserted image from Image class: ");
            builder.InsertImage(image);

            builder.Writeln("\nInserted image from Image class with a custom size: ");
            builder.InsertImage(image, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

            builder.Writeln("\nInserted image from Image class using relative positions: ");
            builder.InsertImage(image, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
                100, 200, 100, WrapType.Square);

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromImageClass.docx");
            //ExEnd
        }

        [Test]
        public void InsertImageFromByteArray()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[])
            //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows different solutions of how to import an image into a document from a byte array.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image image = Image.FromFile(ImageDir + "Logo.jpg");

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, ImageFormat.Png);
                byte[] imageByteArray = ms.ToArray();
 
                builder.Writeln("\nInserted image from byte array: ");
                builder.InsertImage(imageByteArray);

                builder.Writeln("\nInserted image from byte array with a custom size: ");
                builder.InsertImage(imageByteArray, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

                builder.Writeln("\nInserted image from byte array using relative positions: ");
                builder.InsertImage(imageByteArray, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 
                    100, 200, 100, WrapType.Square);
            }

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromByteArray.docx");
            //ExEnd
        }
#else
        [Test]
        public void InsertImageFromImageClassNetStandard2()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows different solutions of how to import an image into a document from Image class (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (SKBitmap bitmap = SKBitmap.Decode(ImageDir + "Logo.jpg"))
            {
                builder.Writeln("\nInserted image from Image class: ");
                builder.InsertImage(bitmap);

                builder.Writeln("\nInserted image from Image class with a custom size: ");
                builder.InsertImage(bitmap, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

                builder.Writeln("\nInserted image from Image class using relative positions: ");
                builder.InsertImage(bitmap, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
                    100, 200, 100, WrapType.Square);
            }

            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromImageClassNetStandard2.docx");
            //ExEnd
        }

        [Test]
        public void InsertImageFromByteArrayNetStandard2()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(Byte[])
            //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
            //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows different solutions of how to import an image into a document from a byte array (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (SKBitmap bitmap = SKBitmap.Decode(ImageDir + "Logo.jpg"))
            {
                using (SKImage image = SKImage.FromBitmap(bitmap))
                {
                    using (SKData data = image.Encode()) // Encode the image (defaults to PNG)
                    {
                        byte[] imageByteArray = data.ToArray();

                        builder.Writeln("\nInserted image from byte array: ");
                        builder.InsertImage(imageByteArray);

                        builder.Writeln("\nInserted image from byte array with a custom size: ");
                        builder.InsertImage(imageByteArray, ConvertUtil.PixelToPoint(250), ConvertUtil.PixelToPoint(144));

                        builder.Writeln("\nInserted image from byte array using relative positions: ");
                        builder.InsertImage(imageByteArray, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 
                            100, 200, 100, WrapType.Square);
                    }
                }
            }
            
            doc.Save(ArtifactsDir + "DocumentBuilderImages.InsertImageFromByteArrayNetStandard2.docx");
            //ExEnd
        }
#endif
    }
}