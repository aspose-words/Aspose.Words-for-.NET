// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing.Drawing2D;
using System.Drawing.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

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

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Emf);
            saveOptions.UseGdiEmfRenderer = false;
            //ExEnd
        }

        [Test]
        public void SaveIntoGif()
        {
            //ExStart
            //ExFor:ImageSaveOptions.UseGdiEmfRenderer
            //ExSummary:Shows how to save specific document page as image file.
            Document doc = new Document(MyDir + "SaveOptions.MyraidPro.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Gif);
            //Define which page will save
            saveOptions.PageIndex = 0;

            doc.Save(MyDir + @"\Artifacts\SaveOptions.MyraidPro Out.gif", saveOptions);
            //ExEnd
        }

        [Test]
        public void QualityOptions()
        {
            //ExStart
            //ExFor:GraphicsQualityOptions
            //ExFor:GraphicsQualityOptions.SmoothingMode
            //ExFor:GraphicsQualityOptions.TextRenderingHint
            //ExSummary:Shows how to set render quality options. 
            Document doc = new Document(MyDir + "SaveOptions.MyraidPro.docx");

            GraphicsQualityOptions qualityOptions = new GraphicsQualityOptions();
            qualityOptions.SmoothingMode = SmoothingMode.AntiAlias;
            qualityOptions.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            saveOptions.GraphicsQualityOptions = qualityOptions;

            doc.Save(MyDir + @"\Artifacts\SaveOptions.QualityOptions Out.jpeg", saveOptions);
            //ExEnd
        }

        [Test]
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
            
            doc.Save(MyDir + @"\Artifacts\ImageSaveOptions.BlackAndWhite Out.png", imageSaveOptions);
            //ExEnd
        }
    }
}