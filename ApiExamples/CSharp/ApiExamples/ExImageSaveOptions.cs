// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;
#if NET462 || JAVA
using System.Drawing.Drawing2D;
using System.Drawing.Text;
#elif NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExImageSaveOptions : ApiExampleBase
    {
        [Test]
        public void OnePage()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:FixedPageSaveOptions
            //ExFor:ImageSaveOptions.PageSet
            //ExSummary:Shows how to render one page from a document to a JPEG image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

            // Set the "PageSet" to "1" to select the second page via
            // the zero-based index to start rendering the document from.
            options.PageSet = new PageSet(1);

            // When we save the document to the JPEG format, Aspose.Words only renders one page.
            // This image will contain one page starting from page two,
            // which will just be the second page of the original document.
            doc.Save(ArtifactsDir + "ImageSaveOptions.OnePage.jpg", options);
            //ExEnd

            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "ImageSaveOptions.OnePage.jpg");
        }

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
                Assert.That(300000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.Renderer.emf").Length));
#elif NETCOREAPP2_1
	            Assert.That(30000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.Renderer.emf").Length));
#endif
            else
                Assert.That(30000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.Renderer.emf").Length));
            //ExEnd

#if NET462 || JAVA
            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "ImageSaveOptions.Renderer.emf");
#endif
        }

        [Test]
        public void PageSet()
        {
            //ExStart
            //ExFor:ImageSaveOptions.PageSet
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
            // We can pass a SaveOptions object to specify a different page to render.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Gif);

            // Render every page of the document to a separate image file.
            for (int i = 1; i <= doc.PageCount; i++)
            {
                saveOptions.PageSet = new PageSet(1);

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

        [Test, Category("SkipMono")]
        public void PageByPage()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:FixedPageSaveOptions
            //ExFor:ImageSaveOptions.PageSet
            //ExSummary:Shows how to render every page of a document to a separate TIFF image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

            for (int i = 0; i < doc.PageCount; i++)
            {
                // Set the "PageSet" property to the number of the first page from
                // which to start rendering the document from.
                options.PageSet = new PageSet(i);

                doc.Save(ArtifactsDir + $"ImageSaveOptions.PageByPage.{i + 1}.tiff", options);
            }
            //ExEnd

            List<string> imageFileNames = Directory.GetFiles(ArtifactsDir, "*.tiff")
                .Where(item => item.Contains("ImageSaveOptions.PageByPage.") && item.EndsWith(".tiff")).ToList();

            Assert.AreEqual(3, imageFileNames.Count);

            foreach (string imageFileName in imageFileNames)
                TestUtil.VerifyImage(816, 1056, imageFileName);
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

            Assert.That(20000, Is.LessThan(new FileInfo(ImageDir + "Logo.jpg").Length));

            // When we save the document as an image, we can pass a SaveOptions object to
            // select a color mode for the image that the saving operation will generate.
            // If we set the "ImageColorMode" property to "ImageColorMode.BlackAndWhite",
            // the saving operation will apply grayscale color reduction while rendering the document.
            // If we set the "ImageColorMode" property to "ImageColorMode.Grayscale", 
            // the saving operation will render the document into a monochrome image.
            // If we set the "ImageColorMode" property to "None", the saving operation will apply the default method
            // and preserve all the document's colors in the output image.
            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.ImageColorMode = imageColorMode;
            
            doc.Save(ArtifactsDir + "ImageSaveOptions.ColorMode.png", imageSaveOptions);

#if NET462 || JAVA
            switch (imageColorMode)
            {
                case ImageColorMode.None:
                    Assert.That(150000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length));
                    break;
                case ImageColorMode.Grayscale:
                    Assert.That(80000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length));
                    break;
                case ImageColorMode.BlackAndWhite:
                    Assert.That(20000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length));
                    break;
            }
#elif NETCOREAPP2_1
            switch (imageColorMode)
            {
                case ImageColorMode.None:
                    Assert.That(120000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length));
                    break;
                case ImageColorMode.Grayscale:
                    Assert.That(80000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length));
                    break;
                case ImageColorMode.BlackAndWhite:
                    Assert.That(20000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length));
                    break;
            }
#endif
            //ExEnd
        }

        [Test]
        public void PaperColor()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.PaperColor
            //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Times New Roman";
            builder.Font.Size = 24;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            builder.InsertImage(ImageDir + "Logo.jpg");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png);

            // Set the "PaperColor" property to a transparent color to apply a transparent
            // background to the document while rendering it to an image.
            imgOptions.PaperColor = Color.Transparent;

            doc.Save(ArtifactsDir + "ImageSaveOptions.PaperColor.Transparent.png", imgOptions);

            // Set the "PaperColor" property to an opaque color to apply that color
            // as the background of the document as we render it to an image.
            imgOptions.PaperColor = Color.LightCoral;

            doc.Save(ArtifactsDir + "ImageSaveOptions.PaperColor.LightCoral.png", imgOptions);
            //ExEnd

            TestUtil.ImageContainsTransparency(ArtifactsDir + "ImageSaveOptions.PaperColor.Transparent.png");
            Assert.Throws<AssertionException>(() =>
                TestUtil.ImageContainsTransparency(ArtifactsDir + "ImageSaveOptions.PaperColor.LightCoral.png"));
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

            Assert.That(20000, Is.LessThan(new FileInfo(ImageDir + "Logo.jpg").Length));

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
                    Assert.That(10000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length));
                    break;
                case ImagePixelFormat.Format16BppRgb555:
                    Assert.That(80000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length));
                    break;
                case ImagePixelFormat.Format24BppRgb:
                    Assert.That(125000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length));
                    break;
                case ImagePixelFormat.Format32BppRgb:
                    Assert.That(150000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length));
                    break;
                case ImagePixelFormat.Format48BppRgb:
                    Assert.That(200000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length));
                    break;
            }
#elif NETCOREAPP2_1
            switch (imagePixelFormat)
            {
                case ImagePixelFormat.Format1bppIndexed:
                    Assert.That(10000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length));
                    break;
                case ImagePixelFormat.Format24BppRgb:
                    Assert.That(70000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length));
                    break;
                case ImagePixelFormat.Format16BppRgb555:
                case ImagePixelFormat.Format32BppRgb:
                case ImagePixelFormat.Format48BppRgb:
                    Assert.That(125000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length));
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
            //ExSummary:Shows how to set the TIFF binarization error threshold when using the Floyd-Steinberg method to render a TIFF image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Hello world!");
            builder.InsertImage(ImageDir + "Logo.jpg");

            // When we save the document as a TIFF, we can pass a SaveOptions object to
            // adjust the dithering that Aspose.Words will apply when rendering this image.
            // The default value of the "ThresholdForFloydSteinbergDithering" property is 128.
            // Higher values tend to produce darker images.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt3,
                TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
                ThresholdForFloydSteinbergDithering = 240
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.FloydSteinbergDithering.tiff", options);
            //ExEnd
            
#if NET462 || JAVA
            TestUtil.VerifyImage(816, 1056, ArtifactsDir + "ImageSaveOptions.FloydSteinbergDithering.tiff");
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
            //ExSummary:Shows how to edit the image while Aspose.Words converts a document to one.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Hello world!");
            builder.InsertImage(ImageDir + "Logo.jpg");

            // When we save the document as an image, we can pass a SaveOptions object to
            // edit the image while the saving operation renders it.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // We can adjust these properties to change the image's brightness and contrast.
                // Both are on a 0-1 scale and are at 0.5 by default.
                ImageBrightness = 0.3f,
                ImageContrast = 0.7f,

                // We can adjust horizontal and vertical resolution with these properties.
                // This will affect the dimensions of the image.
                // The default value for these properties is 96.0, for a resolution of 96dpi.
                HorizontalResolution = 72f,
                VerticalResolution = 72f,

                // We can scale the image using this property. The default value is 1.0, for scaling of 100%.
                // We can use this property to negate any changes in image dimensions that changing the resolution would cause.
                Scale = 96f / 72f
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.EditImage.png", options);
            //ExEnd

            TestUtil.VerifyImage(817, 1057, ArtifactsDir + "ImageSaveOptions.EditImage.png");
        }

        [Test]
        public void JpegQuality()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:FixedPageSaveOptions.JpegQuality
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.#ctor
            //ExFor:ImageSaveOptions.JpegQuality
            //ExSummary:Shows how to configure compression while saving a document as a JPEG.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(ImageDir + "Logo.jpg");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            // Set the "JpegQuality" property to "10" to use stronger compression when rendering the document.
            // This will reduce the file size of the document, but the image will display more prominent compression artifacts.
            imageOptions.JpegQuality = 10;

            doc.Save(ArtifactsDir + "ImageSaveOptions.JpegQuality.HighCompression.jpg", imageOptions);

            Assert.That(20000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.JpegQuality.HighCompression.jpg").Length));

            // Set the "JpegQuality" property to "100" to use weaker compression when rending the document.
            // This will improve the quality of the image at the cost of an increased file size.
            imageOptions.JpegQuality = 100;

            doc.Save(ArtifactsDir + "ImageSaveOptions.JpegQuality.HighQuality.jpg", imageOptions);

            Assert.That(60000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.JpegQuality.HighQuality.jpg").Length));
            //ExEnd
        }

        [Test, Category("SkipMono")]
        public void SaveToTiffDefault()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "ImageSaveOptions.SaveToTiffDefault.tiff");
        }

        [TestCase(TiffCompression.None), Category("SkipMono")]
        [TestCase(TiffCompression.Rle), Category("SkipMono")]
        [TestCase(TiffCompression.Lzw), Category("SkipMono")]
        [TestCase(TiffCompression.Ccitt3), Category("SkipMono")]
        [TestCase(TiffCompression.Ccitt4), Category("SkipMono")]
        public void TiffImageCompression(TiffCompression tiffCompression)
        {
            //ExStart
            //ExFor:TiffCompression
            //ExFor:ImageSaveOptions.TiffCompression
            //ExSummary:Shows how to select the compression scheme to apply to a document that we convert into a TIFF image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImageDir + "Logo.jpg");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

            // Set the "TiffCompression" property to "TiffCompression.None" to apply no compression while saving,
            // which may result in a very large output file.
            // Set the "TiffCompression" property to "TiffCompression.Rle" to apply RLE compression
            // Set the "TiffCompression" property to "TiffCompression.Lzw" to apply LZW compression.
            // Set the "TiffCompression" property to "TiffCompression.Ccitt3" to apply CCITT3 compression.
            // Set the "TiffCompression" property to "TiffCompression.Ccitt4" to apply CCITT4 compression.
            options.TiffCompression = tiffCompression;

            doc.Save(ArtifactsDir + "ImageSaveOptions.TiffImageCompression.tiff", options);

            switch (tiffCompression)
            {
                case TiffCompression.None:
                    Assert.That(3000000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.TiffImageCompression.tiff").Length));
                    break;
                case TiffCompression.Rle:
#if NETCOREAPP2_1
                    Assert.That(6000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.TiffImageCompression.tiff").Length));
#else
                    Assert.That(600000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.TiffImageCompression.tiff").Length));
#endif
                    break;
                case TiffCompression.Lzw:
                    Assert.That(200000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.TiffImageCompression.tiff").Length));
                    break;
                case TiffCompression.Ccitt3:
                    Assert.That(90000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.TiffImageCompression.tiff").Length));
                    break;
                case TiffCompression.Ccitt4:
                    Assert.That(20000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.TiffImageCompression.tiff").Length));
                    break;
            }
            //ExEnd
        }

        [Test]
        public void Resolution()
        {
            //ExStart
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.Resolution
            //ExSummary:Shows how to specify a resolution while rendering a document to PNG.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Times New Roman";
            builder.Font.Size = 24;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            builder.InsertImage(ImageDir + "Logo.jpg");

            // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            // to modify the way in which that method renders the document into an image.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            // Set the "Resolution" property to "72" to render the document in 72dpi.
            options.Resolution = 72;

            doc.Save(ArtifactsDir + "ImageSaveOptions.Resolution.72dpi.png", options);

            Assert.That(120000, Is.AtLeast(new FileInfo(ArtifactsDir + "ImageSaveOptions.Resolution.72dpi.png").Length));

#if NET462 || JAVA
            Image image = Image.FromFile(ArtifactsDir + "ImageSaveOptions.Resolution.72dpi.png");

            Assert.AreEqual(612, image.Width);
            Assert.AreEqual(792, image.Height);
#elif NETCOREAPP2_1 || __MOBILE__
            using (SKBitmap image = SKBitmap.Decode(ArtifactsDir + "ImageSaveOptions.Resolution.72dpi.png")) 
            {
                Assert.AreEqual(612, image.Width);
                Assert.AreEqual(792, image.Height);
            }
#endif
            // Set the "Resolution" property to "300" to render the document in 300dpi.
            options.Resolution = 300;

            doc.Save(ArtifactsDir + "ImageSaveOptions.Resolution.300dpi.png", options);

            Assert.That(700000, Is.LessThan(new FileInfo(ArtifactsDir + "ImageSaveOptions.Resolution.300dpi.png").Length));

#if NET462 || JAVA
            image = Image.FromFile(ArtifactsDir + "ImageSaveOptions.Resolution.300dpi.png");

            Assert.AreEqual(2550, image.Width);
            Assert.AreEqual(3300, image.Height);
#elif NETCOREAPP2_1 || __MOBILE__
            using (SKBitmap image = SKBitmap.Decode(ArtifactsDir + "ImageSaveOptions.Resolution.300dpi.png")) 
            {
                Assert.AreEqual(2550, image.Width);
                Assert.AreEqual(3300, image.Height);
            }
#endif
            //ExEnd
        }

        public void ExportVariousPageRanges()
        {
            //ExStart
            //ExFor:PageSet.#ctor(PageRange[])
            //ExFor:PageRange.#ctor(int, int)
            //ExFor:ImageSaveOptions.PageSet
            //ExSummary:Shows how to extract pages based on exact page ranges.
            Document doc = new Document(MyDir + "Images.docx");

            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Tiff);
            PageSet pageSet = new PageSet(new PageRange(1, 1), new PageRange(2, 3), new PageRange(1, 3), new PageRange(2, 4), new PageRange(1, 1));

            imageOptions.PageSet = pageSet;
            doc.Save(ArtifactsDir + "ImageSaveOptions.ExportVariousPageRanges.tiff", imageOptions);
            //ExEnd
        }
    }
}