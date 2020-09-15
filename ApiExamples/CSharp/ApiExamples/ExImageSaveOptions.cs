// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;
#if NET462 || JAVA
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExImageSaveOptions : ApiExampleBase
    {
        [TestCase(false)]
        [TestCase(true)]
        public void Renderer(bool useGdiEmfRenderer)
        {
            //ExStart
            //ExFor:ImageSaveOptions.UseGdiEmfRenderer
            //ExSummary:Shows how to choose a renderer when converting a document to .emf.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Hello world!");
            builder.InsertImage(ImageDir + "Logo.jpg");

            // When we save the document as an EMF image, we can pass a SaveOptions object to select a renderer for the image.
            // If we set the "UseGdiEmfRenderer" flag to "true", Aspose.Words will use the GDI+ renderer.
            // If we set the "UseGdiEmfRenderer" flag to "false", Aspose.Words will use its own metafile renderer.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Emf);
            saveOptions.UseGdiEmfRenderer = useGdiEmfRenderer;

            doc.Save(ArtifactsDir + "ImageSaveOptions.Renderer.emf", saveOptions);

            // The GDI+ renderer usually creates larger files.
            if (useGdiEmfRenderer)
#if NET462 || JAVA
                Assert.AreEqual(343600, new FileInfo(ArtifactsDir + "ImageSaveOptions.Renderer.emf").Length, 200);
#elif NETCOREAPP2_1
                Assert.AreEqual(21100, new FileInfo(ArtifactsDir + "ImageSaveOptions.Renderer.emf").Length, 200);
#endif
            else
                Assert.AreEqual(21100, new FileInfo(ArtifactsDir + "ImageSaveOptions.Renderer.emf").Length, 200);
            //ExEnd

#if NET462 || JAVA
            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "ImageSaveOptions.Renderer.emf");
#endif
        }

        [Test]
        public void PageIndex()
        {
            //ExStart
            //ExFor:ImageSaveOptions.PageIndex
            //ExSummary:Shows how to specify which page in a document to render as an image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Hello world! This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 3.");

            Assert.AreEqual(3, doc.PageCount);

            // When we save the document as an image, Aspose.Words only renders the first page by default.
            // We can pass a SaveOptions object to specify a different to page to render.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Gif);

            // Render every page of the document to a separate image file.
            for (int i = 1; i <= doc.PageCount; i++)
            {
                saveOptions.PageIndex = 1;

                doc.Save(ArtifactsDir + $"ImageSaveOptions.PageIndex.Page {i}.gif", saveOptions);
            }
            //ExEnd

            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "ImageSaveOptions.PageIndex.Page 1.gif");
            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "ImageSaveOptions.PageIndex.Page 2.gif");
            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "ImageSaveOptions.PageIndex.Page 3.gif");
            Assert.False(File.Exists(ArtifactsDir + "ImageSaveOptions.PageIndex.Page 4.gif"));
        }

#if NET462 || JAVA
        [Test]
        public void GraphicsQuality()
        {
            //ExStart
            //ExFor:GraphicsQualityOptions
            //ExFor:GraphicsQualityOptions.CompositingMode
            //ExFor:GraphicsQualityOptions.CompositingQuality
            //ExFor:GraphicsQualityOptions.InterpolationMode
            //ExFor:GraphicsQualityOptions.StringFormat
            //ExFor:GraphicsQualityOptions.SmoothingMode
            //ExFor:GraphicsQualityOptions.TextRenderingHint
            //ExFor:ImageSaveOptions.GraphicsQualityOptions
            //ExSummary:Shows how to set render quality options while converting documents to image formats. 
            Document doc = new Document(MyDir + "Rendering.docx");

            GraphicsQualityOptions qualityOptions = new GraphicsQualityOptions
            {
                SmoothingMode = SmoothingMode.AntiAlias,
                TextRenderingHint = TextRenderingHint.ClearTypeGridFit,
                CompositingMode = CompositingMode.SourceOver,
                CompositingQuality = CompositingQuality.HighQuality,
                InterpolationMode = InterpolationMode.High,
                StringFormat = StringFormat.GenericTypographic
            };

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            saveOptions.GraphicsQualityOptions = qualityOptions;

            doc.Save(ArtifactsDir + "ImageSaveOptions.GraphicsQuality.jpg", saveOptions);
            //ExEnd

            TestUtil.VerifyImage(794, 1122, ArtifactsDir + "ImageSaveOptions.GraphicsQuality.jpg");
        }

        [TestCase(MetafileRenderingMode.Vector), Category("SkipMono")]
        [TestCase(MetafileRenderingMode.Bitmap), Category("SkipMono")]
        [TestCase(MetafileRenderingMode.VectorWithFallback), Category("SkipMono")]
        public void WindowsMetaFile(MetafileRenderingMode metafileRenderingMode)
        {
            //ExStart
            //ExFor:ImageSaveOptions.MetafileRenderingOptions
            //ExSummary:Shows how to set the rendering mode when saving documents with Windows Metafile images to other image formats. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertImage(Image.FromFile(ImageDir + "Windows MetaFile.wmf"));
            
            // When we save the document as an image, we can pass a SaveOptions object to
            // determine how the saving operation will process Windows Metafiles in the document.
            // If we set the "RenderingMode" property to "MetafileRenderingMode.Vector",
            // or "MetafileRenderingMode.VectorWithFallback", we will render all metafiles as vector graphics.
            // If we set the "RenderingMode" property to "MetafileRenderingMode.Bitmap", we will render all metafiles as bitmaps.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.MetafileRenderingOptions.RenderingMode = metafileRenderingMode;
            
            doc.Save(ArtifactsDir + "ImageSaveOptions.WindowsMetaFile.png", options);
            //ExEnd

            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "ImageSaveOptions.WindowsMetaFile.png");
        }
#endif

        [TestCase(ImageColorMode.BlackAndWhite)]
        [TestCase(ImageColorMode.Grayscale)]
        [TestCase(ImageColorMode.None)]
        public void ColorMode(ImageColorMode imageColorMode)
        {
            //ExStart
            //ExFor:ImageColorMode
            //ExFor:ImageSaveOptions.ImageColorMode
            //ExSummary:Shows how to set a color mode when rendering documents.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Hello world!");
            builder.InsertImage(ImageDir + "Logo.jpg");

            Assert.AreEqual(20100, new FileInfo(ImageDir + "Logo.jpg").Length, 200);

            // When we save the document as an image, we can pass a SaveOptions object to
            // select a color mode for the image that the saving operation will generate.
            // If we set the "ImageColorMode" property to "ImageColorMode.BlackAndWhite",
            // the saving operation will apply grayscale color reduction while rendering the document
            // so it only consists of black and white.
            // If we set the "ImageColorMode" property to "ImageColorMode.Grayscale", 
            // the saving operation will render the document into a monochrome image.
            // If we set the "ImageColorMode" property to "None", the saving operation will apply the default method
            // and preserve all the colors of the document in the output image.
            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.ImageColorMode = imageColorMode;
            
            doc.Save(ArtifactsDir + "ImageSaveOptions.ColorMode.png", imageSaveOptions);

#if NET462 || JAVA
            switch (imageColorMode)
            {
                case ImageColorMode.None:
                    Assert.AreEqual(174100, new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length, 200);
                    break;
                case ImageColorMode.Grayscale:
                    Assert.AreEqual(89100, new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length, 200);
                    break;
                case ImageColorMode.BlackAndWhite:
                    Assert.AreEqual(14800, new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length, 200);
                    break;
            }
#elif NETCOREAPP2_1
            switch (imageColorMode)
            {
                case ImageColorMode.None:
                    Assert.AreEqual(131700, new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length, 200);
                    break;
                case ImageColorMode.Grayscale:
                    Assert.AreEqual(89100, new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length, 200);
                    break;
                case ImageColorMode.BlackAndWhite:
                    Assert.AreEqual(10900, new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length, 200);
                    break;
            }
#endif
            //ExEnd
        }

        [TestCase(ImagePixelFormat.Format1bppIndexed)]
        [TestCase(ImagePixelFormat.Format16BppRgb555)]
        [TestCase(ImagePixelFormat.Format24BppRgb)]
        [TestCase(ImagePixelFormat.Format32BppRgb)]
        [TestCase(ImagePixelFormat.Format48BppRgb)]
        public void PixelFormat(ImagePixelFormat imagePixelFormat)
        {
            //ExStart
            //ExFor:ImagePixelFormat
            //ExFor:ImageSaveOptions.Clone
            //ExFor:ImageSaveOptions.PixelFormat
            //ExSummary:Shows how to select a bit-per-pixel rate with which to render a document to an image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Hello world!");
            builder.InsertImage(ImageDir + "Logo.jpg");

            Assert.AreEqual(20100, new FileInfo(ImageDir + "Logo.jpg").Length, 200);

            // When we save the document as an image, we can pass a SaveOptions object to
            // select a pixel format for the image that the saving operation will generate.
            // Various bit per pixel rates will affect the quality and file size of the generated image.
            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PixelFormat = imagePixelFormat;

            // We can clone ImageSaveOptions instances.
            Assert.AreNotEqual(imageSaveOptions, imageSaveOptions.Clone());

            doc.Save(ArtifactsDir + "ImageSaveOptions.PixelFormat.png", imageSaveOptions);

#if NET462 || JAVA
            switch (imagePixelFormat)
            {
                case ImagePixelFormat.Format1bppIndexed:
                    Assert.AreEqual(2300, new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length, 200);
                    break;
                case ImagePixelFormat.Format16BppRgb555:
                    Assert.AreEqual(87600, new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length, 200);
                    break;
                case ImagePixelFormat.Format24BppRgb:
                    Assert.AreEqual(158200, new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length, 200);
                    break;
                case ImagePixelFormat.Format32BppRgb:
                    Assert.AreEqual(174100, new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length, 200);
                    break;
                case ImagePixelFormat.Format48BppRgb:
                    Assert.AreEqual(211200, new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length, 200);
                    break;
            }
#elif NETCOREAPP2_1
            switch (imagePixelFormat)
            {
                case ImagePixelFormat.Format1bppIndexed:
                    Assert.AreEqual(5600, new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length, 200);
                    break;
                case ImagePixelFormat.Format24BppRgb:
                    Assert.AreEqual(76000, new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length, 200);
                    break;
                case ImagePixelFormat.Format16BppRgb555:
                case ImagePixelFormat.Format32BppRgb:
                case ImagePixelFormat.Format48BppRgb:
                    Assert.AreEqual(131700, new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length, 200);
                    break;
            }
#endif
            //ExEnd
        }
        
        [Test, Category("SkipMono")]
        public void FloydSteinbergDithering()
        {
            //ExStart
            //ExFor:ImageBinarizationMethod
            //ExFor:ImageSaveOptions.ThresholdForFloydSteinbergDithering
            //ExFor:ImageSaveOptions.TiffBinarizationMethod
            //ExSummary: Shows how to control the threshold for TIFF binarization in the Floyd-Steinberg method
            Document doc = new Document (MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt3,
                ImageColorMode = ImageColorMode.Grayscale,
                TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
                // The default value of this property is 128. The higher value, the darker image
                ThresholdForFloydSteinbergDithering = 254
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.FloydSteinbergDithering.tiff", options);
            //ExEnd
            
#if NET462 || JAVA
            TestUtil.VerifyImage(794, 1123, ArtifactsDir + "ImageSaveOptions.FloydSteinbergDithering.tiff");
#endif
        }

        [Test]
        public void EditImage()
        {
            //ExStart
            //ExFor:ImageSaveOptions.HorizontalResolution
            //ExFor:ImageSaveOptions.ImageBrightness
            //ExFor:ImageSaveOptions.ImageContrast
            //ExFor:ImageSaveOptions.SaveFormat
            //ExFor:ImageSaveOptions.Scale
            //ExFor:ImageSaveOptions.VerticalResolution
            //ExSummary:Shows how to edit image.
            Document doc = new Document(MyDir + "Rendering.docx");

            // When saving the document as an image, we can use an ImageSaveOptions object to edit various aspects of it
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                ImageBrightness = 0.3f,     // 0 - 1 scale, default at 0.5
                ImageContrast = 0.7f,       // 0 - 1 scale, default at 0.5
                HorizontalResolution = 72f, // Default at 96.0 meaning 96dpi, image dimensions will be affected if we change resolution
                VerticalResolution = 72f,   // Default at 96.0 meaning 96dpi
                Scale = 96f / 72f           // Default at 1.0 for normal scale, can be used to negate resolution impact in image size
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.EditImage.png", options);
            //ExEnd

            TestUtil.VerifyImage(794, 1123, ArtifactsDir + "ImageSaveOptions.EditImage.png");
        }
    }
}