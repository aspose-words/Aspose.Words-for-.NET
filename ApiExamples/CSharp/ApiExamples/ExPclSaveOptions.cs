// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExPclSaveOptions : ApiExampleBase
    {
        [Test]
        public void RasterizeElements()
        {
            //ExStart
            //ExFor:PclSaveOptions
            //ExFor:PclSaveOptions.RasterizeTransformedElements
            //ExSummary:Shows how rasterized or not transformed elements before saving.
            Document doc = new Document(MyDir + "Document.EpubConversion.doc");

            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.RasterizeTransformedElements = true;

            doc.Save(MyDir + @"\Artifacts\Document.EpubConversion.pcl", saveOptions);
            //ExEnd
        }

        [Test]
        public void SetPrinterFont()
        {
            //ExStart
            //ExFor:PclSaveOptions.AddPrinterFont(string, string)
            //ExFor:PclSaveOptions.FallbackFontName
            //ExSummary:Shows how to add information about font that is uploaded to the printer and set the font that will be used if no expected font is found in printer and built-in fonts collections.
            Document doc = new Document(MyDir + "Document.EpubConversion.doc");
            
            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.AddPrinterFont("Courier", "Courier");
            saveOptions.FallbackFontName = "Times New Roman";

            doc.Save(MyDir + @"\Artifacts\Document.EpubConversion.pcl", saveOptions);
            //ExEnd
        }
    }
}