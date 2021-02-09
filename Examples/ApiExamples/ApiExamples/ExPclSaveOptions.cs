// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
            //ExFor:PclSaveOptions.SaveFormat
            //ExFor:PclSaveOptions.RasterizeTransformedElements
            //ExSummary:Shows how to rasterize complex elements while saving a document to PCL.
            Document doc = new Document(MyDir + "Rendering.docx");

            PclSaveOptions saveOptions = new PclSaveOptions
            {
                SaveFormat = SaveFormat.Pcl,
                RasterizeTransformedElements = true
            };

            doc.Save(ArtifactsDir + "PclSaveOptions.RasterizeElements.pcl", saveOptions);
            //ExEnd
        }

        [Test]
        public void FallbackFontName()
        {
            //ExStart
            //ExFor:PclSaveOptions.FallbackFontName
            //ExSummary:Shows how to declare a font that a printer will apply to printed text as a substitute should its original font be unavailable.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Non-existent font";
            builder.Write("Hello world!");

            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.FallbackFontName = "Times New Roman";
            
            // This document will instruct the printer to apply "Times New Roman" to the text with the missing font.
            // Should "Times New Roman" also be unavailable, the printer will default to the "Arial" font.
            doc.Save(ArtifactsDir + "PclSaveOptions.SetPrinterFont.pcl", saveOptions);
            //ExEnd
        }

        [Test]
        public void AddPrinterFont()
        {
            //ExStart
            //ExFor:PclSaveOptions.AddPrinterFont(string, string)
            //ExSummary:Shows how to get a printer to substitute all instances of a specific font with a different font. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Courier";
            builder.Write("Hello world!");

            PclSaveOptions saveOptions = new PclSaveOptions();
            saveOptions.AddPrinterFont("Courier New", "Courier");

            // When printing this document, the printer will use the "Courier New" font
            // to access places where our document used the "Courier" font.
            doc.Save(ArtifactsDir + "PclSaveOptions.AddPrinterFont.pcl", saveOptions);
            //ExEnd
        }

        [Test]
        [Description("This test is a manual check that PaperTray information is preserved in the output pcl document.")]
        public void GetPreservedPaperTrayInformation()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            // Paper tray information is now preserved when saving document to PCL format.
            // Following information is transferred from document's model to PCL file.
            foreach (Section section in doc.Sections.OfType<Section>())
            {
                section.PageSetup.FirstPageTray = 15;
                section.PageSetup.OtherPagesTray = 12;
            }

            doc.Save(ArtifactsDir + "PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
        }
    }
}