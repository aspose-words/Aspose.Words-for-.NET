// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
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

        //ToDo: Need to check gold test
        [Test]
        public void SaveIntoGif()
        {
            //ExStart
            //ExFor:ImageSaveOptions.UseGdiEmfRenderer
            //ExSummary:Shows how to save specific document page as gif.
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
            //ExFor:ImageSaveOptions.GraphicsQualityOptions
            //ExFor:GraphicsQualityOptions
            //ExSummary:Shows how to use rendering hints
            Document doc = new Document(MyDir + "SaveOptions.MyraidPro.docx");

            GraphicsQualityOptions qualityOptions = new GraphicsQualityOptions();
            qualityOptions.SmoothingMode = SmoothingMode.AntiAlias;
            qualityOptions.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            saveOptions.GraphicsQualityOptions = qualityOptions;

            doc.Save(MyDir + @"\Artifacts\SaveOptions.QualityOptions Out.jpeg", saveOptions);
            //ExEnd
        }
    }
}