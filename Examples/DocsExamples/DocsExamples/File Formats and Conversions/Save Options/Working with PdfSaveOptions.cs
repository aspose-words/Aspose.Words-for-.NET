using System;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
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

        //ExStart:RenderMetafileToBitmap
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
        //ExEnd:RenderMetafileToBitmap
        //ExEnd:PdfRenderWarnings

        [Test]
        public void DigitallySignedPdfUsingCertificateHolder()
        {
            //ExStart:DigitallySignedPdfUsingCertificateHolder
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
            //ExStart:EmbeddAllFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will be embedded with all fonts found in the document.
            PdfSaveOptions saveOptions = new PdfSaveOptions { EmbedFullFonts = true };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EmbeddedFontsInPdf.pdf", saveOptions);
            //ExEnd:EmbeddAllFonts
        }

        [Test]
        public void EmbeddedSubsetFonts()
        {
            //ExStart:EmbeddSubsetFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will contain subsets of the fonts in the document.
            // Only the glyphs used in the document are included in the PDF fonts.
            PdfSaveOptions saveOptions = new PdfSaveOptions { EmbedFullFonts = false };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EmbeddSubsetFonts.pdf", saveOptions);
            //ExEnd:EmbeddSubsetFonts
        }

        [Test]
        public void DisableEmbedWindowsFonts()
        {
            //ExStart:DisableEmbedWindowsFonts
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
            Document doc = new Document(MyDir + "Bookmarks in headers and footers.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;
            saveOptions.HeaderFooterBookmarksExportMode = HeaderFooterBookmarksExportMode.First;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf", saveOptions);
            //ExEnd:ExportHeaderFooterBookmarks
        }

        [Test]
        public void ScaleWmfFontsToMetafileSize()
        {
            //ExStart:ScaleWmfFontsToMetafileSize
            Document doc = new Document(MyDir + "WMF with text.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions
            {
                ScaleWmfFontsToMetafileSize = false
            };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
            // then Aspose.Words renders this metafile to a bitmap.
            PdfSaveOptions saveOptions = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ScaleWmfFontsToMetafileSize.pdf", saveOptions);
            //ExEnd:ScaleWmfFontsToMetafileSize
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
            //ExStart:ConversionToPDF17
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { Compliance = PdfCompliance.Pdf17 };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ConversionToPdf17.pdf", saveOptions);
            //ExEnd:ConversionToPDF17
        }

        [Test]
        public void DownsamplingImages()
        {
            //ExStart:DownsamplingImages
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
        public void SetOutlineOptions()
        {
            //ExStart:SetOutlineOptions
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 3;
            saveOptions.OutlineOptions.ExpandedOutlineLevels = 1;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.SetOutlineOptions.pdf", saveOptions);
            //ExEnd:SetOutlineOptions
        }

        [Test]
        public void CustomPropertiesExport()
        {
            //ExStart:CustomPropertiesExport
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
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro, while editing the .pdf.
            PdfSaveOptions saveOptions = new PdfSaveOptions { ExportDocumentStructure = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ExportDocumentStructure.pdf", saveOptions);
            //ExEnd:ExportDocumentStructure
        }

        [Test]
        public void PdfImageComppression()
        {
            //ExStart:PdfImageComppression
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg, PreserveFormFields = true
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.PdfImageCompression.pdf", saveOptions);

            PdfSaveOptions saveOptionsA1B = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b,
                ImageCompression = PdfImageCompression.Jpeg,
                JpegQuality = 100, // Use JPEG compression at 50% quality to reduce file size.
                ImageColorSpaceExportMode = PdfImageColorSpaceExportMode.SimpleCmyk
            };
            

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.PdfImageCompression.Pdf_A1b.pdf", saveOptionsA1B);
            //ExEnd:PdfImageComppression
        }

        [Test]
        public void UpdateLastPrintedProperty()
        {
            //ExStart:UpdateIfLastPrinted
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions { UpdateLastPrintedProperty = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.UpdateIfLastPrinted.pdf", saveOptions);
            //ExEnd:UpdateIfLastPrinted
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
    }
}