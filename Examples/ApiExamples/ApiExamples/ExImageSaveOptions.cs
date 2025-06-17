// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using NUnit.Framework;
using Aspose.Words.Drawing;
#if NET461_OR_GREATER || JAVA
using System.Drawing.Drawing2D;
using System.Drawing.Text;
#elif NET5_0_OR_GREATER || __MOBILE__
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
            //ExFor:PageSet
            //ExFor:PageSet.#ctor(Int32)
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
            //ExEnd
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

            Assert.That(doc.PageCount, Is.EqualTo(3));

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
            Assert.That(File.Exists(ArtifactsDir + "ImageSaveOptions.PageIndex.Page 4.gif"), Is.False);
        }

#if NET461_OR_GREATER || JAVA
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

        [Test]
        public void UseTileFlipMode()
        {
            //ExStart
            //ExFor:GraphicsQualityOptions.UseTileFlipMode
            //ExSummary:Shows how to prevent the white line appears when rendering with a high resolution.
            Document doc = new Document(MyDir + "Shape high dpi.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            ShapeRenderer renderer = shape.GetShapeRenderer();

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 500, GraphicsQualityOptions = new GraphicsQualityOptions { UseTileFlipMode = true }
            };
            renderer.Save(ArtifactsDir + "ImageSaveOptions.UseTileFlipMode.png", saveOptions);
            //ExEnd
        }
#endif

        [TestCase(MetafileRenderingMode.Vector), Category("SkipMono")]
        [TestCase(MetafileRenderingMode.Bitmap), Category("SkipMono")]
        [TestCase(MetafileRenderingMode.VectorWithFallback), Category("SkipMono")]
        public void WindowsMetaFile(MetafileRenderingMode metafileRenderingMode)
        {
            //ExStart
            //ExFor:ImageSaveOptions.MetafileRenderingOptions
            //ExFor:MetafileRenderingOptions.UseGdiRasterOperationsEmulation
            //ExSummary:Shows how to set the rendering mode when saving documents with Windows Metafile images to other image formats. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImageDir + "Windows MetaFile.wmf");

            // When we save the document as an image, we can pass a SaveOptions object to
            // determine how the saving operation will process Windows Metafiles in the document.
            // If we set the "RenderingMode" property to "MetafileRenderingMode.Vector",
            // or "MetafileRenderingMode.VectorWithFallback", we will render all metafiles as vector graphics.
            // If we set the "RenderingMode" property to "MetafileRenderingMode.Bitmap", we will render all metafiles as bitmaps.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.MetafileRenderingOptions.RenderingMode = metafileRenderingMode;
            // Aspose.Words uses GDI+ for raster operations emulation, when value is set to true.
            options.MetafileRenderingOptions.UseGdiRasterOperationsEmulation = true;

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
            //ExFor:ImageSaveOptions.ImageSize
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
                // Export page at 2325x5325 pixels and 600 dpi.
                options.Resolution = 600;
                options.ImageSize = new Size(2325, 5325);

                doc.Save(ArtifactsDir + $"ImageSaveOptions.PageByPage.{i + 1}.tiff", options);
            }
            //ExEnd

            List<string> imageFileNames = Directory.GetFiles(ArtifactsDir, "*.tiff")
                .Where(item => item.Contains("ImageSaveOptions.PageByPage.") && item.EndsWith(".tiff")).ToList();
            Assert.That(imageFileNames.Count, Is.EqualTo(3));
        }

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
            //ExEnd

            var testedImageLength = new FileInfo(ArtifactsDir + "ImageSaveOptions.ColorMode.png").Length;

#if NET461_OR_GREATER || JAVA
            switch (imageColorMode)
            {
                case ImageColorMode.None:
                    Assert.That(testedImageLength < 175000, Is.True);
                    break;
                case ImageColorMode.Grayscale:
                    Assert.That(testedImageLength < 90000, Is.True);
                    break;
                case ImageColorMode.BlackAndWhite:
                    Assert.That(testedImageLength < 15000, Is.True);
                    break;
            }
#elif NET5_0_OR_GREATER
            switch (imageColorMode)
            {
                case ImageColorMode.None:
                    Assert.That(testedImageLength < 132000, Is.True);
                    break;
                case ImageColorMode.Grayscale:
                    Assert.That(testedImageLength < 90000, Is.True);
                    break;
                case ImageColorMode.BlackAndWhite:
                    Assert.That(testedImageLength < 11000, Is.True);
                    break;
            }
#endif            
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
        [TestCase(ImagePixelFormat.Format16BppRgb565)]
        [TestCase(ImagePixelFormat.Format24BppRgb)]
        [TestCase(ImagePixelFormat.Format32BppRgb)]
        [TestCase(ImagePixelFormat.Format32BppArgb)]
        [TestCase(ImagePixelFormat.Format32BppPArgb)]
        [TestCase(ImagePixelFormat.Format48BppRgb)]
        [TestCase(ImagePixelFormat.Format64BppArgb)]
        [TestCase(ImagePixelFormat.Format64BppPArgb)]
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

            // When we save the document as an image, we can pass a SaveOptions object to
            // select a pixel format for the image that the saving operation will generate.
            // Various bit per pixel rates will affect the quality and file size of the generated image.
            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PixelFormat = imagePixelFormat;

            // We can clone ImageSaveOptions instances.
            Assert.That(imageSaveOptions.Clone(), Is.Not.EqualTo(imageSaveOptions));

            doc.Save(ArtifactsDir + "ImageSaveOptions.PixelFormat.png", imageSaveOptions);
            //ExEnd

            var testedImageLength = new FileInfo(ArtifactsDir + "ImageSaveOptions.PixelFormat.png").Length;

#if NET461_OR_GREATER || JAVA
            switch (imagePixelFormat)
            {
                case ImagePixelFormat.Format1bppIndexed:
                    Assert.That(testedImageLength < 2500, Is.True);
                    break;
                case ImagePixelFormat.Format16BppRgb565:
                    Assert.That(testedImageLength < 104000, Is.True);
                    break;
                case ImagePixelFormat.Format16BppRgb555:
                    Assert.That(testedImageLength < 88000, Is.True);
                    break;
                case ImagePixelFormat.Format24BppRgb:
                    Assert.That(testedImageLength < 160000, Is.True);
                    break;
                case ImagePixelFormat.Format32BppRgb:
                case ImagePixelFormat.Format32BppArgb:
                    Assert.That(testedImageLength < 175000, Is.True);
                    break;
                case ImagePixelFormat.Format48BppRgb:
                    Assert.That(testedImageLength < 212000, Is.True);
                    break;
                case ImagePixelFormat.Format64BppArgb:
                case ImagePixelFormat.Format64BppPArgb:
                    Assert.That(testedImageLength < 239000, Is.True);
                    break;
            }
#elif NET5_0_OR_GREATER
            switch (imagePixelFormat)
            {
                case ImagePixelFormat.Format1bppIndexed:
                    Assert.That(testedImageLength < 7500, Is.True);
                    break;
                case ImagePixelFormat.Format24BppRgb:
                    Assert.That(testedImageLength < 77000, Is.True);
                    break;
                case ImagePixelFormat.Format16BppRgb565:
                case ImagePixelFormat.Format16BppRgb555:
                case ImagePixelFormat.Format32BppRgb:
                case ImagePixelFormat.Format32BppArgb:
                case ImagePixelFormat.Format48BppRgb:
                case ImagePixelFormat.Format64BppArgb:
                case ImagePixelFormat.Format64BppPArgb:
                    Assert.That(testedImageLength < 132000, Is.True);
                    break;
            }
#endif
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

            List<string> imageFileNames = Directory.GetFiles(ArtifactsDir, "*.tiff")
                .Where(item => item.Contains("ImageSaveOptions.FloydSteinbergDithering.") && item.EndsWith(".tiff")).ToList();
            Assert.That(imageFileNames.Count, Is.EqualTo(1));
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

            // Set the "JpegQuality" property to "100" to use weaker compression when rending the document.
            // This will improve the quality of the image at the cost of an increased file size.
            imageOptions.JpegQuality = 100;
            doc.Save(ArtifactsDir + "ImageSaveOptions.JpegQuality.HighQuality.jpg", imageOptions);
            //ExEnd

            Assert.That(new FileInfo(ArtifactsDir + "ImageSaveOptions.JpegQuality.HighCompression.jpg").Length < 18000, Is.True);
            Assert.That(new FileInfo(ArtifactsDir + "ImageSaveOptions.JpegQuality.HighQuality.jpg").Length < 75000, Is.True);
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
            //ExEnd

            var testedImageLength = new FileInfo(ArtifactsDir + "ImageSaveOptions.TiffImageCompression.tiff").Length;

            switch (tiffCompression)
            {
                case TiffCompression.None:
                    Assert.That(testedImageLength < 3450000, Is.True);
                    break;
                case TiffCompression.Rle:
#if NET5_0_OR_GREATER
                    Assert.That(testedImageLength < 7500, Is.True);
#else
                    Assert.That(testedImageLength < 687000, Is.True);
#endif
                    break;
                case TiffCompression.Lzw:
                    Assert.That(testedImageLength < 250000, Is.True);
                    break;
                case TiffCompression.Ccitt3:
#if NET5_0_OR_GREATER
                    Assert.That(testedImageLength < 6100, Is.True);
#else
                    Assert.That(testedImageLength < 8300, Is.True);
#endif
                    break;
                case TiffCompression.Ccitt4:
                    Assert.That(testedImageLength < 1700, Is.True);
                    break;
            }
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

            // Set the "Resolution" property to "300" to render the document in 300dpi.
            options.Resolution = 300;
            doc.Save(ArtifactsDir + "ImageSaveOptions.Resolution.300dpi.png", options);
            //ExEnd

            TestUtil.VerifyImage(612, 792, ArtifactsDir + "ImageSaveOptions.Resolution.72dpi.png");
            TestUtil.VerifyImage(2550, 3300, ArtifactsDir + "ImageSaveOptions.Resolution.300dpi.png");
        }

        [Test]
        public void ExportVariousPageRanges()
        {
            //ExStart
            //ExFor:PageSet.#ctor(PageRange[])
            //ExFor:PageRange
            //ExFor:PageRange.#ctor(int, int)
            //ExFor:ImageSaveOptions.PageSet
            //ExSummary:Shows how to extract pages based on exact page ranges.
            Document doc = new Document(MyDir + "Images.docx");

            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Tiff);
            PageSet pageSet = new PageSet(new PageRange(1, 1), new PageRange(2, 3), new PageRange(1, 3),
                new PageRange(2, 4), new PageRange(1, 1));

            imageOptions.PageSet = pageSet;
            doc.Save(ArtifactsDir + "ImageSaveOptions.ExportVariousPageRanges.tiff", imageOptions);
            //ExEnd
        }

        [Test]
        public void RenderInkObject()
        {
            //ExStart
            //ExFor:SaveOptions.ImlRenderingMode
            //ExFor:ImlRenderingMode
            //ExSummary:Shows how to render Ink object.
            Document doc = new Document(MyDir + "Ink object.docx");

            // Set 'ImlRenderingMode.InkML' ignores fall-back shape of ink (InkML) object and renders InkML itself.
            // If the rendering result is unsatisfactory,
            // please use 'ImlRenderingMode.Fallback' to get a result similar to previous versions.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                ImlRenderingMode = ImlRenderingMode.InkML
            };

            doc.Save(ArtifactsDir + "ImageSaveOptions.RenderInkObject.jpeg", saveOptions);
            //ExEnd
        }

        [Test]
        public void GridLayout()
        {
            //ExStart:GridLayout
            //GistId:70330eacdfc2e253f00a9adea8972975
            //ExFor:ImageSaveOptions.PageLayout
            //ExFor:MultiPageLayout
            //ExSummary:Shows how to save the document into JPG image with multi-page layout settings.
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);
            // Set up a grid layout with:
            // - 3 columns per row.
            // - 10pts spacing between pages (horizontal and vertical).
            options.PageLayout = MultiPageLayout.Grid(3, 10, 10);

            // Alternative layouts:
            // options.PageLayout = MultiPageLayout.Horizontal(10);
            // options.PageLayout = MultiPageLayout.Vertical(10);

            // Customize the background and border.
            options.PageLayout.BackColor = Color.LightGray;
            options.PageLayout.BorderColor = Color.Blue;
            options.PageLayout.BorderWidth = 2;

            doc.Save(ArtifactsDir + "ImageSaveOptions.GridLayout.jpg", options);
            //ExEnd:GridLayout
        }
    }
}
