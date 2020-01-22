// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;
using Document = Aspose.Words.Document;
using IWarningCallback = Aspose.Words.IWarningCallback;
using PdfSaveOptions = Aspose.Words.Saving.PdfSaveOptions;
using SaveFormat = Aspose.Words.SaveFormat;
using SaveOptions = Aspose.Words.Saving.SaveOptions;
using WarningInfo = Aspose.Words.WarningInfo;
using WarningType = Aspose.Words.WarningType;
#if NETFRAMEWORK
using Image = System.Drawing.Image;
#else 
using SkiaSharp;
#endif
#if NETFRAMEWORK || NETSTANDARD2_0
using Aspose.Pdf.Facades;
using Aspose.Pdf.Annotations;
#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExPdfSaveOptions : ApiExampleBase
    {
        [Test]
        public void CreateMissingOutlineLevels()
        {
            //ExStart
            //ExFor:OutlineOptions.CreateMissingOutlineLevels
            //ExFor:ParagraphFormat.IsHeading
            //ExFor:PdfSaveOptions.OutlineOptions
            //ExFor:PdfSaveOptions.SaveFormat
            //ExSummary:Shows how to create missing outline levels saving the document in PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Creating TOC entries
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            Assert.True(builder.ParagraphFormat.IsHeading);

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;

            builder.Writeln("Heading 1.1.1.1");
            builder.Writeln("Heading 1.1.1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading9;

            builder.Writeln("Heading 1.1.1.1.1.1.1.1.1");
            builder.Writeln("Heading 1.1.1.1.1.1.1.1.2");

            // Create "PdfSaveOptions" with some mandatory parameters
            // "HeadingsOutlineLevels" specifies how many levels of headings to include in the document outline
            // "CreateMissingOutlineLevels" determining whether or not to create missing heading levels
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9;
            pdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels = true;
            pdfSaveOptions.SaveFormat = SaveFormat.Pdf;

            doc.Save(ArtifactsDir + "CreateMissingOutlineLevels.pdf", pdfSaveOptions);
            //ExEnd

            #if NETFRAMEWORK || NETSTANDARD2_0
            // Bind PDF with Aspose.PDF
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "CreateMissingOutlineLevels.pdf");

            // Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(11, bookmarks.Count);
            #endif
        }

        [Test]
        [Category("SkipMono")]
        public void WithoutUpdateFields()
        {
            //ExStart
            //ExFor:PdfSaveOptions.Clone
            //ExFor:SaveOptions.UpdateFields
            //ExSummary:Shows how to update fields before saving into a PDF document.
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                UpdateFields = false
            };

            // PdfSaveOptions objects can be cloned
            Assert.AreNotSame(pdfSaveOptions, pdfSaveOptions.Clone());

            doc.Save(ArtifactsDir + "UpdateFields_False.pdf", pdfSaveOptions);
            //ExEnd

            #if NETFRAMEWORK || NETSTANDARD2_0
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "UpdateFields_False.pdf");

            // Get text fragment by search String
            Aspose.Pdf.Text.TextFragmentAbsorber textFragmentAbsorber = new Aspose.Pdf.Text.TextFragmentAbsorber("Page  of");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            // Assert that fields are not updated
            Assert.AreEqual("Page  of", textFragmentAbsorber.TextFragments[1].Text);
            #endif
        }

        [Test]
        [Category("SkipMono")]
        public void WithUpdateFields()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { UpdateFields = true };

            doc.Save(ArtifactsDir + "UpdateFields_False.pdf", pdfSaveOptions);

            #if NETFRAMEWORK || NETSTANDARD2_0
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "UpdateFields_False.pdf");

            // Get text fragment by search String from PDF document
            Aspose.Pdf.Text.TextFragmentAbsorber textFragmentAbsorber = new Aspose.Pdf.Text.TextFragmentAbsorber("Page 1 of 2");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            // Assert that fields are updated
            Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
            #endif
        }

        // For assert this test you need to open "SaveOptions.PdfImageCompression PDF_A_1_B Out.pdf" and "SaveOptions.PdfImageCompression PDF_A_1_A Out.pdf" 
        // and check that header image in this documents are equal header image in the "SaveOptions.PdfImageComppression Out.pdf" 
        [Test]
        public void ImageCompression()
        {
            //ExStart
            //ExFor:PdfSaveOptions.Compliance
            //ExFor:PdfSaveOptions.ImageCompression
            //ExFor:PdfSaveOptions.ImageColorSpaceExportMode
            //ExFor:PdfSaveOptions.JpegQuality
            //ExFor:PdfImageCompression
            //ExFor:PdfCompliance
            //ExFor:PdfImageColorSpaceExportMode
            //ExSummary:Shows how to save images to PDF using JPEG encoding to decrease file size.
            Document doc = new Document(MyDir + "SaveOptions.PdfImageCompression.rtf");
            
            PdfSaveOptions options = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg,
                PreserveFormFields = true
            };
            doc.Save(ArtifactsDir + "SaveOptions.PdfImageCompression.pdf", options);

            PdfSaveOptions optionsA1B = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b,
                ImageCompression = PdfImageCompression.Jpeg,
                JpegQuality = 100, // Use JPEG compression at 50% quality to reduce file size
                ImageColorSpaceExportMode = PdfImageColorSpaceExportMode.SimpleCmyk
            };

            doc.Save(ArtifactsDir + "SaveOptions.PdfImageComppression PDF_A_1_B.pdf", optionsA1B);        
            //ExEnd

            PdfSaveOptions optionsA1A = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a,
                ExportDocumentStructure = true,
                ImageCompression = PdfImageCompression.Jpeg
            };

            doc.Save(ArtifactsDir + "SaveOptions.PdfImageComppression PDF_A_1_A.pdf", optionsA1A);
        }

        [Test]
        public void ColorRendering()
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:FixedPageSaveOptions.ColorMode
            //ExSummary:Shows how change image color with save options property
            // Open document with color image
            Document doc = new Document(MyDir + "Rendering.doc");
            // Set grayscale mode for document
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { ColorMode = ColorMode.Grayscale };
            
            // Assert that color image in document was grey
            doc.Save(ArtifactsDir + "ColorMode.PdfGrayscaleMode.pdf", pdfSaveOptions);
            //ExEnd
        }

        [Test]
        public void WindowsBarPdfTitle()
        {
            //ExStart
            //ExFor:PdfSaveOptions.DisplayDocTitle
            //ExSummary:Shows how to display title of the document as title bar.
            Document doc = new Document(MyDir + "Rendering.doc");
            doc.BuiltInDocumentProperties.Title = "Windows bar pdf title";
            
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { DisplayDocTitle = true };

            doc.Save(ArtifactsDir + "PdfTitle.pdf", pdfSaveOptions);
            //ExEnd

            #if NETFRAMEWORK || NETSTANDARD2_0
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfTitle.pdf");

            Assert.IsTrue(pdfDocument.DisplayDocTitle);
            Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
            #endif
        }

        [Test]
        public void MemoryOptimization()
        {
            //ExStart
            //ExFor:SaveOptions.CreateSaveOptions(SaveFormat)
            //ExFor:SaveOptions.MemoryOptimization
            //ExSummary:Shows an option to optimize memory consumption when you work with large documents.
            Document doc = new Document(MyDir + "SaveOptions.MemoryOptimization.doc");
            
            // When set to true it will improve document memory footprint but will add extra time to processing
            // This optimization is only applied during save operation
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
            saveOptions.MemoryOptimization = true;

            doc.Save(ArtifactsDir + "SaveOptions.MemoryOptimization.pdf", saveOptions);
            //ExEnd
        }

        [Test]
        [TestCase(@"https://www.google.com/search?q= aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", true)]
        [TestCase(@"https://www.google.com/search?q=%20aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", true)]
        [TestCase(@"https://www.google.com/search?q= aspose", "app.launchURL(\"https://www.google.com/search?q= aspose\", true);", false)]
        [TestCase(@"https://www.google.com/search?q=%20aspose", "app.launchURL(\"https://www.google.com/search?q=%20aspose\", true);", false)]
        public void EscapeUri(string uri, string result, bool isEscaped)
        {
            //ExStart
            //ExFor:PdfSaveOptions.EscapeUri
            //ExFor:PdfSaveOptions.OpenHyperlinksInNewWindow
            //ExSummary: Shows how to escape hyperlinks or not in the document.
            DocumentBuilder builder = new DocumentBuilder();
            builder.InsertHyperlink("Testlink", uri, false);

            // Set this property to false if you are sure that hyperlinks in document's model are already escaped
            PdfSaveOptions options = new PdfSaveOptions();
            options.EscapeUri = isEscaped;
            options.OpenHyperlinksInNewWindow = true;

            builder.Document.Save(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf", options);
            //ExEnd

            #if NETFRAMEWORK || NETSTANDARD2_0
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf");

            // Get first page
            Aspose.Pdf.Page page = pdfDocument.Pages[1];
            // Get the first link annotation
            LinkAnnotation linkAnnot = (LinkAnnotation) page.Annotations[1];

            JavascriptAction action = (JavascriptAction) linkAnnot.Action;
            string uriText = action.Script;

            Assert.AreEqual(result, uriText);
            #endif
        }

        [Test]
        [Category("SkipMono")]
        public void HandleBinaryRasterWarnings()
        {
            //ExStart
            //ExFor:MetafileRenderingMode
            //ExFor:MetafileRenderingOptions
            //ExFor:MetafileRenderingOptions.EmulateRasterOperations
            //ExFor:MetafileRenderingOptions.RenderingMode
            //ExFor:IWarningCallback
            //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
            //ExSummary:Shows added fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
            Document doc = new Document(MyDir + "PdfSaveOptions.HandleRasterWarnings.doc");

            MetafileRenderingOptions metafileRenderingOptions =
                new MetafileRenderingOptions
                {
                    EmulateRasterOperations = false,
                    RenderingMode = MetafileRenderingMode.VectorWithFallback
                };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = metafileRenderingOptions;

            doc.Save(ArtifactsDir + "PdfSaveOptions.HandleRasterWarnings.pdf", saveOptions);

            Assert.AreEqual(1, callback.Warnings.Count);
            Assert.True(callback.Warnings[0].Description.Contains("R2_XORPEN"));
        }

        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // For now type of warnings about unsupported metafile records changed from
                // DataLoss/UnexpectedContent to MinorFormattingLoss
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    Warnings.Warning(info);
                }
            }

            public WarningInfoCollection Warnings = new WarningInfoCollection();
        }
        //ExEnd

        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.None)]
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.First)]
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.All)]
        public void HeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode headerFooterBookmarksExportMode)
        {
            //ExStart
            //ExFor:HeaderFooterBookmarksExportMode
            //ExFor:PdfSaveOptions.HeaderFooterBookmarksExportMode
            //ExFor:OutlineOptions
            //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
            //ExSummary:Shows how bookmarks in headers/footers are exported to pdf.
            Document doc = new Document(MyDir + "PdfSaveOption.HeaderFooterBookmarksExportMode.docx");

            // You can specify how bookmarks in headers/footers are exported
            // There is a several options for this:
            // "None" - Bookmarks in headers/footers are not exported
            // "First" - Only bookmark in first header/footer of the section is exported
            // "All" - Bookmarks in all headers/footers are exported
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                HeaderFooterBookmarksExportMode = headerFooterBookmarksExportMode,
                OutlineOptions = { DefaultBookmarksOutlineLevel = 1 }
            };
            doc.Save(ArtifactsDir + "PdfSaveOption.HeaderFooterBookmarksExportMode.pdf", saveOptions);
            //ExEnd
        }

        [Test]
        public void UnsupportedImageFormatWarning()
        {
            Document doc = new Document(MyDir + "PdfSaveOptions.TestCorruptedImage.docx");

            SaveWarningCallback saveWarningCallback = new SaveWarningCallback();
            doc.WarningCallback = saveWarningCallback;

            doc.Save(ArtifactsDir + "PdfSaveOption.HeaderFooterBookmarksExportMode.pdf", SaveFormat.Pdf);

            Assert.That(saveWarningCallback.SaveWarnings[0].Description,
                Is.EqualTo("Image can not be processed. Possibly unsupported image format."));
        }

        public class SaveWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine($"{info.WarningType}: {info.Description}.");
                    SaveWarnings.Warning(info);
                }
            }

            internal WarningInfoCollection SaveWarnings = new WarningInfoCollection();
		}
		
		[Test]
        public void FontsScaledToMetafileSize()
        {
            //ExStart
            //ExFor:MetafileRenderingOptions.ScaleWmfFontsToMetafileSize
            //ExSummary:Shows how to WMF fonts scaling according to metafile size on the page.
            Document doc = new Document(MyDir + "PdfSaveOptions.FontsScaledToMetafileSize.docx");

            // There is a several options for this:
            // 'True' - Aspose.Words emulates font scaling according to metafile size on the page
            // 'False' - Aspose.Words displays the fonts as metafile is rendered to its default size
            // Use 'False' option is used only when metafile is rendered as vector graphics
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions.ScaleWmfFontsToMetafileSize = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.FontsScaledToMetafileSize.pdf", saveOptions);
            //ExEnd
        }

        [Test]
        public void AdditionalTextPositioning()
        {
            //ExStart
            //ExFor:PdfSaveOptions.AdditionalTextPositioning
            //ExSummary:Show how to write additional text positioning operators.
            Document doc = new Document(MyDir + "PdfSaveOptions.AdditionalTextPositioning.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            // This may help to overcome issues with inaccurate text positioning with some printers
            saveOptions.AdditionalTextPositioning = true;
            saveOptions.TextCompression = PdfTextCompression.None;

            doc.Save(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
            //ExEnd
        }

        [Test]
        public void SaveAsPdfBookFold()
        {
            //ExStart
            //ExFor:PdfSaveOptions.UseBookFoldPrintingSettings
            //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
            // Open a document with multiple paragraphs
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Configure both page setup and PdfSaveOptions to create a book fold
            foreach (Section s in doc.Sections)
            {
                s.PageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;
            }

            PdfSaveOptions options = new PdfSaveOptions();
            options.UseBookFoldPrintingSettings = true;

            // In order to make a booklet, we will need to print this document, stack the pages
            // in the order they come out of the printer and then fold down the middle
            doc.Save(ArtifactsDir + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
            //ExEnd
        }

        [Test]
        public void ZoomBehaviour()
        {
            //ExStart
            //ExFor:PdfSaveOptions.PageMode
            //ExFor:PdfSaveOptions.ZoomBehavior
            //ExFor:PdfSaveOptions.ZoomFactor
            //ExFor:PdfPageMode
            //ExFor:PdfZoomBehavior
            //ExSummary:Shows how to set the default zooming of an output PDF to 1/4 of default size.
            // Open a document with multiple paragraphs
            Document doc = new Document(MyDir + "Rendering.doc");

            PdfSaveOptions options = new PdfSaveOptions();
            options.ZoomBehavior = PdfZoomBehavior.ZoomFactor;
            options.ZoomFactor = 25;
            options.PageMode = PdfPageMode.UseThumbs;

            // When opening the .pdf with a viewer such as Adobe Acrobat Pro, the zoom level will be at 25% by default,
            // with thumbnails for each page to the left
            doc.Save(ArtifactsDir + "PdfSaveOptions.ZoomBehaviour.pdf", options);
            //ExEnd
        }

        [Test]
        public void NoteHyperlinks()
        {
            //ExStart
            //ExFor:PdfSaveOptions.CreateNoteHyperlinks
            //ExSummary:Shows how to make footnotes and endnotes work like hyperlinks.
            // Open a document with footnotes/endnotes
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            // Creating a PdfSaveOptions instance with this flag set will convert footnote/endnote number symbols in the text
            // into hyperlinks pointing to the footnotes, and the actual footnotes/endnotes at the end of pages into links to their
            // referenced body text
            PdfSaveOptions options = new PdfSaveOptions();
            options.CreateNoteHyperlinks = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf", options);
            //ExEnd
        }

        [Test]
        public void CustomPropertiesExport()
        {
            //ExStart
            //ExFor:PdfCustomPropertiesExport
            //ExFor:PdfSaveOptions.CustomPropertiesExport
            //ExFor:SaveOptions.DmlEffectsRenderingMode
            //ExSummary:Shows how to export custom properties while saving to .pdf.
            Document doc = new Document();

            // Add a custom document property that doesn't use the name of some built in properties
            doc.CustomDocumentProperties.Add("Company", "My value");
            
            // Configure the PdfSaveOptions like this will display the properties
            // in the "Document Properties" menu of Adobe Acrobat Pro
            PdfSaveOptions options = new PdfSaveOptions();
            options.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

            doc.Save(ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
            //ExEnd
        }

        [Test]
        public void DrawingML()
        {
            //ExStart
            //ExFor:DmlRenderingMode
            //ExFor:PdfSaveOptions.DmlEffectsRenderingMode
            //ExFor:SaveOptions.DmlRenderingMode
            //ExSummary:Shows how to configure DrawingML rendering quality with PdfSaveOptions.
            Document doc = new Document(MyDir + "DrawingMLEffects.docx");

            // Creating a new PdfSaveOptions object and setting its DmlEffectsRenderingMode to "None" will
            // strip the shapes of all their shading effects in the output pdf
            PdfSaveOptions options = new PdfSaveOptions();
            options.DmlEffectsRenderingMode = DmlEffectsRenderingMode.None;

            doc.Save(ArtifactsDir + "PdfSaveOptions.DrawingML.pdf", options);
            //ExEnd
        }

        [Test]
        public void ExportDocumentStructure()
        {
            //ExStart
            //ExFor:PdfSaveOptions.ExportDocumentStructure
            //ExSummary:Shows how to convert a .docx to .pdf while preserving the document structure.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro, while editing the .pdf
            PdfSaveOptions options = new PdfSaveOptions();
            options.ExportDocumentStructure = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
            //ExEnd
        }

        #if NETFRAMEWORK
        [Test]
        public void PreblendImages()
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreblendImages
            //ExSummary:Shows how to preblend images with transparent backgrounds.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image img = Image.FromFile(ImageDir + "TransparentBG.png");
            builder.InsertImage(img);

            // Create a PdfSaveOptions object and setting this flag may change the quality and size of the output .pdf
            // because of the way some images are rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.PreblendImages = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PreblendImages.pdf", options);
            //ExEnd
        }
        #else
        [Test]
        public void PreblendImagesNetStandard2()
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreblendImages
            //ExSummary:Shows how to preblend images with transparent backgrounds (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (SKBitmap image = SKBitmap.Decode(ImageDir + "TransparentBG.png"))
            {
                builder.InsertImage(image);
            }

            // Create a PdfSaveOptions object and setting this flag may change the quality and size of the output .pdf
            // because of the way some images are rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.PreblendImages = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PreblendImages.pdf", options);
            //ExEnd
        }
        #endif

        [Test]
        public void PdfDigitalSignature()
        {
            //ExStart
            //ExFor:PdfDigitalSignatureDetails
            //ExFor:PdfDigitalSignatureDetails.#ctor
            //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
            //ExFor:PdfDigitalSignatureDetails.HashAlgorithm
            //ExFor:PdfDigitalSignatureDetails.Location
            //ExFor:PdfDigitalSignatureDetails.Reason
            //ExFor:PdfDigitalSignatureDetails.SignatureDate
            //ExFor:PdfDigitalSignatureHashAlgorithm
            //ExFor:PdfSaveOptions.DigitalSignatureDetails
            //ExSummary:Shows how to sign a generated PDF using Aspose.Words.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Signed PDF contents.");

            // Load the certificate from disk
            // The other constructor overloads can be used to load certificates from different locations
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            // Pass the certificate and details to the save options class to sign with
            PdfSaveOptions options = new PdfSaveOptions();
            DateTime signingTime = DateTime.Now;
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", signingTime);

            // We can use this attribute to set a different hash algorithm
            options.DigitalSignatureDetails.HashAlgorithm = PdfDigitalSignatureHashAlgorithm.Sha256;

            Assert.AreEqual("Test Signing", options.DigitalSignatureDetails.Reason);
            Assert.AreEqual("Aspose Office", options.DigitalSignatureDetails.Location);
            Assert.AreEqual(signingTime.ToUniversalTime(), options.DigitalSignatureDetails.SignatureDate);

            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf");
            //ExEnd
        }

        [Test]
        public void PdfDigitalSignatureTimestamp()
        {
            //ExStart
            //ExFor:PdfDigitalSignatureDetails.TimestampSettings
            //ExFor:PdfDigitalSignatureTimestampSettings
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String)
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String,TimeSpan)
            //ExFor:PdfDigitalSignatureTimestampSettings.Password
            //ExFor:PdfDigitalSignatureTimestampSettings.ServerUrl
            //ExFor:PdfDigitalSignatureTimestampSettings.Timeout
            //ExFor:PdfDigitalSignatureTimestampSettings.UserName
            //ExSummary:Shows how to sign a generated PDF and timestamp it using Aspose.Words.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Signed PDF contents.");

            // Create a digital signature for the document that we will save
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", DateTime.Now);

            // We can set a verified timestamp for our signature as well, with a valid timestamp authority
            options.DigitalSignatureDetails.TimestampSettings =
                new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword");

            // The default lifespan of the timestamp is 100 seconds
            Assert.AreEqual(100.0d, options.DigitalSignatureDetails.TimestampSettings.Timeout.TotalSeconds);

            // We can set our own timeout period via the constructor
            options.DigitalSignatureDetails.TimestampSettings =
                new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", TimeSpan.FromMinutes(30));

            Assert.AreEqual(1800.0d, options.DigitalSignatureDetails.TimestampSettings.Timeout.TotalSeconds);
            Assert.AreEqual("https://freetsa.org/tsr", options.DigitalSignatureDetails.TimestampSettings.ServerUrl);
            Assert.AreEqual("JohnDoe", options.DigitalSignatureDetails.TimestampSettings.UserName);
            Assert.AreEqual("MyPassword", options.DigitalSignatureDetails.TimestampSettings.Password);

            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");
            //ExEnd
        }
    }
}