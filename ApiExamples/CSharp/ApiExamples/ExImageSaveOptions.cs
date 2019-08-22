// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

#if !(NETSTANDARD2_0 || __MOBILE__)
using System.Drawing.Drawing2D;
using System.Drawing.Text;

#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExImageSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseGdiEmfRenderer()
        {
            //ExStart
            //ExFor:ImageSaveOptions.UseGdiEmfRenderer
            //ExSummary:Shows how to save metafiles directly without using GDI+ to EMF.
            Document doc = new Document(MyDir + "SaveOptions.MyraidPro.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Emf)
            {
                UseGdiEmfRenderer = false
            };

            doc.Save(ArtifactsDir + "SaveOptions.UseGdiEmfRenderer.docx", saveOptions);
            //ExEnd
        }

        [Test]
        public void SaveIntoGif()
        {
            //ExStart
            //ExFor:ImageSaveOptions.PageIndex
            //ExSummary:Shows how to save specific document page as image file.
            Document doc = new Document(MyDir + "SaveOptions.MyraidPro.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Gif)
            {
                PageIndex = 0 // Define which page will save
            };

            doc.Save(ArtifactsDir + "SaveOptions.MyraidPro.gif", saveOptions);
            //ExEnd
        }

#if !(NETSTANDARD2_0 || __MOBILE__)
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
            //ExSummary:Shows how to set render quality options when converting documents to image formats. 
            Document doc = new Document(MyDir + "SaveOptions.MyriadPro.docx");

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

            doc.Save(ArtifactsDir + "SaveOptions.GraphicsQuality.jpeg", saveOptions);
            //ExEnd
        }
#endif

        [Test]
        [Category("SkipMono")]
        public void ConverImageColorsToBlackAndWhite()
        {
            //ExStart
            //ExFor:ImageSaveOptions.ImageColorMode
            //ExFor:ImageSaveOptions.PixelFormat
            //ExSummary:Show how to convert document images to black and white with 1 bit per pixel
            Document doc = new Document(MyDir + "ImageSaveOptions.BlackAndWhite.docx");

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.ImageColorMode = ImageColorMode.BlackAndWhite;
            imageSaveOptions.PixelFormat = ImagePixelFormat.Format1bppIndexed;
            
            doc.Save(ArtifactsDir + "ImageSaveOptions.BlackAndWhite.png", imageSaveOptions);
            //ExEnd
        }

        [Test]
        public void ThresholdForFloydSteinbergDithering()
        {
            //ExStart
            //ExFor:ImageSaveOptions.ThresholdForFloydSteinbergDithering
            //ExSummary: Shows how to control the threshold for TIFF binarization in the Floyd-Steinberg method
            Document doc = new Document (MyDir + "ImagesSaveOptions.ThresholdForFloydSteinbergDithering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt3,
                ImageColorMode = ImageColorMode.Grayscale,
                TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
                ThresholdForFloydSteinbergDithering = 254 // The default value of this property is 128. The higher value, the darker image.
            };

            doc.Save(ArtifactsDir + "ImagesSaveOptions.ThresholdForFloydSteinbergDithering.tiff", options);
            //ExEnd
        }
    }
}