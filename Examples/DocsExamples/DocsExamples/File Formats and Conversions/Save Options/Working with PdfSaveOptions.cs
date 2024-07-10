using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithPdfSaveOptions : DocsExamplesBase
    {
        [Test]
        public void DisplayDocTitleInWindowTitlebar()
        {
            //ExStart:DisplayDocTitleInWindowTitlebar
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { DisplayDocTitle = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
            //ExEnd:DisplayDocTitleInWindowTitlebar
        }

        [Test]
        //ExStart:PdfRenderWarnings
        //GistId:f9c5250f94e595ea3590b3be679475ba
        public void PdfRenderWarnings()
        {
            Document doc = new Document(MyDir + "WMF with image.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions
            {
                EmulateRasterOperations = false, RenderingMode = MetafileRenderingMode.VectorWithFallback
            };

            PdfSaveOptions saveOptions = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };

            // If Aspose.Words cannot correctly render some of the metafile records
            // to vector graphics then Aspose.Words renders this metafile to a bitmap.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.PdfRenderWarnings.pdf", saveOptions);

            // While the file saves successfully, rendering warnings that occurred during saving are collected here.
            foreach (WarningInfo warningInfo in callback.mWarnings)
            {
                Console.WriteLine(warningInfo.Description);
            }
        }

        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during
            /// document load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // For now type of warnings about unsupported metafile records changed
                // from DataLoss/UnexpectedContent to MinorFormattingLoss.
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    mWarnings.Warning(info);
                }
            }

            public WarningInfoCollection mWarnings = new WarningInfoCollection();
        }
        //ExEnd:PdfRenderWarnings

        [Test]
        public void DigitallySignedPdfUsingCertificateHolder()
        {
            //ExStart:DigitallySignedPdfUsingCertificateHolder
            //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("Test Signed PDF.");

            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                DigitalSignatureDetails = new PdfDigitalSignatureDetails(
                    CertificateHolder.Create(MyDir + "morzal.pfx", "aw"), "reason", "location",
                    DateTime.Now)
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.DigitallySignedPdfUsingCertificateHolder.pdf", saveOptions);
            //ExEnd:DigitallySignedPdfUsingCertificateHolder
        }

        [Test]
        public void EmbeddedAllFonts()
        {
            //ExStart:EmbeddedAllFonts
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will be embedded with all fonts found in the document.
            PdfSaveOptions saveOptions = new PdfSaveOptions { EmbedFullFonts = true };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EmbeddedAllFonts.pdf", saveOptions);
            //ExEnd:EmbeddedAllFonts
        }

        [Test]
        public void EmbeddedSubsetFonts()
        {
            //ExStart:EmbeddedSubsetFonts
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will contain subsets of the fonts in the document.
            // Only the glyphs used in the document are included in the PDF fonts.
            PdfSaveOptions saveOptions = new PdfSaveOptions { EmbedFullFonts = false };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EmbeddedSubsetFonts.pdf", saveOptions);
            //ExEnd:EmbeddedSubsetFonts
        }

        [Test]
        public void DisableEmbedWindowsFonts()
        {
            //ExStart:DisableEmbedWindowsFonts
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will be saved without embedding standard windows fonts.
            PdfSaveOptions saveOptions = new PdfSaveOptions { FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNone };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.DisableEmbedWindowsFonts.pdf", saveOptions);
            //ExEnd:DisableEmbedWindowsFonts
        }

        [Test]
        public void SkipEmbeddedArialAndTimesRomanFonts()
        {
            //ExStart:SkipEmbeddedArialAndTimesRomanFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.SkipEmbeddedArialAndTimesRomanFonts.pdf", saveOptions);
            //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
        }

        [Test]
        public void AvoidEmbeddingCoreFonts()
        {
            //ExStart:AvoidEmbeddingCoreFonts
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
            PdfSaveOptions saveOptions = new PdfSaveOptions { UseCoreFonts = true };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.AvoidEmbeddingCoreFonts.pdf", saveOptions);
            //ExEnd:AvoidEmbeddingCoreFonts
        }
        
        [Test]
        public void EscapeUri()
        {
            //ExStart:EscapeUri
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertHyperlink("Testlink", 
                "https://www.google.com/search?q=%2Fthe%20test", false);
            builder.Writeln();
            builder.InsertHyperlink("https://www.google.com/search?q=%2Fthe%20test", 
                "https://www.google.com/search?q=%2Fthe%20test", false);

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EscapeUri.pdf");
            //ExEnd:EscapeUri
        }

        [Test]
        public void ExportHeaderFooterBookmarks()
        {
            //ExStart:ExportHeaderFooterBookmarks
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Bookmarks in headers and footers.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;
            saveOptions.HeaderFooterBookmarksExportMode = HeaderFooterBookmarksExportMode.First;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf", saveOptions);
            //ExEnd:ExportHeaderFooterBookmarks
        }

        [Test]
        public void EmulateRenderingToSizeOnPage()
        {
            //ExStart:EmulateRenderingToSizeOnPage
            Document doc = new Document(MyDir + "WMF with text.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions
            {
                EmulateRenderingToSizeOnPage = false
            };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
            // then Aspose.Words renders this metafile to a bitmap.
            PdfSaveOptions saveOptions = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EmulateRenderingToSizeOnPage.pdf", saveOptions);
            //ExEnd:EmulateRenderingToSizeOnPage
        }

        [Test]
        public void AdditionalTextPositioning()
        {
            //ExStart:AdditionalTextPositioning
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { AdditionalTextPositioning = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
            //ExEnd:AdditionalTextPositioning
        }

        [Test]
        public void ConversionToPdf17()
        {
            //ExStart:ConversionToPdf17
            //GistId:a53bdaad548845275c1b9556ee21ae65
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { Compliance = PdfCompliance.Pdf17 };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ConversionToPdf17.pdf", saveOptions);
            //ExEnd:ConversionToPdf17
        }

        [Test]
        public void DownsamplingImages()
        {
            //ExStart:DownsamplingImages
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Rendering.docx");

            // We can set a minimum threshold for downsampling.
            // This value will prevent the second image in the input document from being downsampled.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                DownsampleOptions = { Resolution = 36, ResolutionThreshold = 128 }
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.DownsamplingImages.pdf", saveOptions);
            //ExEnd:DownsamplingImages
        }

        [Test]
        public void OutlineOptions()
        {
            //ExStart:OutlineOptions
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 3;
            saveOptions.OutlineOptions.ExpandedOutlineLevels = 1;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.OutlineOptions.pdf", saveOptions);
            //ExEnd:OutlineOptions
        }

        [Test]
        public void CustomPropertiesExport()
        {
            //ExStart:CustomPropertiesExport
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document();
            doc.CustomDocumentProperties.Add("Company", "Aspose");

            PdfSaveOptions saveOptions = new PdfSaveOptions { CustomPropertiesExport = PdfCustomPropertiesExport.Standard };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.CustomPropertiesExport.pdf", saveOptions);
            //ExEnd:CustomPropertiesExport
        }

        [Test]
        public void ExportDocumentStructure()
        {
            //ExStart:ExportDocumentStructure
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro, while editing the .pdf.
            PdfSaveOptions saveOptions = new PdfSaveOptions { ExportDocumentStructure = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ExportDocumentStructure.pdf", saveOptions);
            //ExEnd:ExportDocumentStructure
        }

        [Test]
        public void ImageCompression()
        {
            //ExStart:ImageCompression
            //GistId:6debb84fc15c7e5b8e35384d9c116215
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg, PreserveFormFields = true
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ImageCompression.pdf", saveOptions);

            PdfSaveOptions saveOptionsA2U = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u,
                ImageCompression = PdfImageCompression.Jpeg,
                JpegQuality = 100, // Use JPEG compression at 50% quality to reduce file size.
            };

            

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ImageCompression_A2u.pdf", saveOptionsA2U);
            //ExEnd:ImageCompression
        }

        [Test]
        public void UpdateLastPrinted()
        {
            //ExStart:UpdateLastPrinted
            //GistId:83e5c469d0e72b5114fb8a05a1d01977
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { UpdateLastPrintedProperty = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.UpdateLastPrinted.pdf", saveOptions);
            //ExEnd:UpdateLastPrinted
        }

        [Test]
        public void Dml3DEffectsRendering()
        {
            //ExStart:Dml3DEffectsRendering
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.Dml3DEffectsRendering.pdf", saveOptions);
            //ExEnd:Dml3DEffectsRendering
        }

        [Test]
        public void InterpolateImages()
        {
            //ExStart:SetImageInterpolation
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { InterpolateImages = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.InterpolateImages.pdf", saveOptions);
            //ExEnd:SetImageInterpolation
        }

        [Test]
        public void OptimizeOutput()
        {
            //ExStart:OptimizeOutput
            //GistId:a53bdaad548845275c1b9556ee21ae65
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { OptimizeOutput = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.OptimizeOutput.pdf", saveOptions);
            //ExEnd:OptimizeOutput
        }

        [Test]
        public void UpdateScreenTip()
        {
            //ExStart:UpdateScreenTip
            //GistId:8b0ab362f95040ada1255a0473acefe2
            Document doc = new Document(MyDir + "Table of contents.docx");

            var tocHyperLinks = doc.Range.Fields
                .Where(f => f.Type == FieldType.FieldHyperlink)
                .Cast<FieldHyperlink>()
                .Where(f => f.SubAddress.StartsWith("#_Toc"));

            foreach (FieldHyperlink link in tocHyperLinks)
                link.ScreenTip = link.DisplayResult;

            PdfSaveOptions saveOptions = new PdfSaveOptions()
            {
                Compliance = PdfCompliance.PdfUa1,
                DisplayDocTitle = true,
                ExportDocumentStructure = true,
            };
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 3;
            saveOptions.OutlineOptions.CreateMissingOutlineLevels = true;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.UpdateScreenTip.pdf", saveOptions);
            //ExEnd:UpdateScreenTip
        }
    }
}