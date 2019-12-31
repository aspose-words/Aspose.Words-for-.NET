// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Linq;
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

            PclSaveOptions saveOptions = new PclSaveOptions
            {
                RasterizeTransformedElements = true
            };

            doc.Save(ArtifactsDir + "PclSaveOptions.RasterizeElements.pcl", saveOptions);
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

            doc.Save(ArtifactsDir + "PclSaveOptions.SetPrinterFont.pcl", saveOptions);
            //ExEnd
        }

        [Test]
        [Ignore("This test is manual check that PaperTray information are preserved in pcl document.")]
        public void GetPreservedPaperTrayInformation()
        {
            Document doc = new Document(MyDir + "Document.EpubConversion.doc");

            // Paper tray information is now preserved when saving document to PCL format
            // Following information is transferred from document's model to PCL file
            foreach (Section section in doc.Sections.OfType<Section>())
            {
                section.PageSetup.FirstPageTray = 15;
                section.PageSetup.OtherPagesTray = 12;
            }

            doc.Save(ArtifactsDir + "PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
        }
    }
}