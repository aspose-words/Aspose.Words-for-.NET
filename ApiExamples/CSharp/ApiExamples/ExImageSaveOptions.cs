// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
        public void QualityOptions()
        {
            //ExStart
            //ExFor:GraphicsQualityOptions
            //ExFor:GraphicsQualityOptions.SmoothingMode
            //ExFor:GraphicsQualityOptions.TextRenderingHint
            //ExSummary:Shows how to set render quality options. 
            Document doc = new Document(MyDir + "SaveOptions.MyraidPro.docx");

            GraphicsQualityOptions qualityOptions = new GraphicsQualityOptions
            {
                SmoothingMode = SmoothingMode.AntiAlias,
                TextRenderingHint = TextRenderingHint.ClearTypeGridFit
            };

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            saveOptions.GraphicsQualityOptions = qualityOptions;

            doc.Save(ArtifactsDir + "SaveOptions.QualityOptions.jpeg", saveOptions);
            //ExEnd
        }
#endif

        [Test]
        [Platform(Exclude = "Linux")]
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
    }
}