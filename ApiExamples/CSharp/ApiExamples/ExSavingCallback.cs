// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExSavingCallback : ApiExampleBase
    {
        [Test]
        public void CheckThatAllMethodsArePresent()
        {
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
            htmlFixedSaveOptions.PageSavingCallback = new CustomFileNamePageSavingCallback();

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PageSavingCallback = new CustomFileNamePageSavingCallback();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.PageSavingCallback = new CustomFileNamePageSavingCallback();

            PsSaveOptions psSaveOptions = new PsSaveOptions();
            psSaveOptions.PageSavingCallback = new CustomFileNamePageSavingCallback();

            SvgSaveOptions svgSaveOptions = new SvgSaveOptions();
            svgSaveOptions.PageSavingCallback = new CustomFileNamePageSavingCallback();

            XamlFixedSaveOptions xamlFixedSaveOptions = new XamlFixedSaveOptions();
            xamlFixedSaveOptions.PageSavingCallback = new CustomFileNamePageSavingCallback();

            XpsSaveOptions xpsSaveOptions = new XpsSaveOptions();
            xpsSaveOptions.PageSavingCallback = new CustomFileNamePageSavingCallback();
        }

        //ExStart
        //ExFor:IPageSavingCallback
        //ExFor:IPageSavingCallback.PageSaving(PageSavingArgs)
        //ExFor:PageSavingArgs
        //ExFor:PageSavingArgs.PageFileName
        //ExFor:PageSavingArgs.KeepPageStreamOpen
        //ExFor:PageSavingArgs.PageIndex
        //ExFor:PageSavingArgs.PageStream
        //ExFor:FixedPageSaveOptions.PageSavingCallback
        //ExSummary:Shows how to use a callback to save a document to HTML page by page.
        [Test] //ExSkip
        public void PageFileNames()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we convert the document to HTML.
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();

            // We will save each page in this document to a separate HTML file in the local file system.
            // Set a callback that allows us to name each output HTML document.
            htmlFixedSaveOptions.PageSavingCallback = new CustomFileNamePageSavingCallback();

            doc.Save(ArtifactsDir + "SavingCallback.PageFileNames.html", htmlFixedSaveOptions);

            string[] filePaths = Directory.GetFiles(ArtifactsDir).Where(
                s => s.StartsWith(ArtifactsDir + "SavingCallback.PageFileNames.Page_")).OrderBy(s => s).ToArray();

            Assert.AreEqual(3, filePaths.Length);
        }

        /// <summary>
        /// Saves all pages to a file and directory specified within.
        /// </summary>
        private class CustomFileNamePageSavingCallback : IPageSavingCallback
        {
            public void PageSaving(PageSavingArgs args)
            {
                string outFileName = $"{ArtifactsDir}SavingCallback.PageFileNames.Page_{args.PageIndex}.html";

                // Below are two ways of specifying where Aspose.Words will save each page of the document.
                // 1 -  Set a filename for the output page file:
                args.PageFileName = outFileName;

                // 2 -  Create a custom stream for the output page file:
                args.PageStream = new FileStream(outFileName, FileMode.Create);

                Assert.False(args.KeepPageStreamOpen);
            }
        }
        //ExEnd

        //ExStart
        //ExFor:DocumentPartSavingArgs
        //ExFor:DocumentPartSavingArgs.Document
        //ExFor:DocumentPartSavingArgs.DocumentPartFileName
        //ExFor:DocumentPartSavingArgs.DocumentPartStream
        //ExFor:DocumentPartSavingArgs.KeepDocumentPartStreamOpen
        //ExFor:IDocumentPartSavingCallback
        //ExFor:IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs)
        //ExFor:IImageSavingCallback
        //ExFor:IImageSavingCallback.ImageSaving
        //ExFor:ImageSavingArgs
        //ExFor:ImageSavingArgs.ImageFileName
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.DocumentPartSavingCallback
        //ExFor:HtmlSaveOptions.ImageSavingCallback
        //ExSummary:Shows how to split a document into parts and save them.
        [Test] //ExSkip
        public void DocumentPartsFileNames()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            string outFileName = "SavingCallback.DocumentPartsFileNames.html";

            // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we convert the document to HTML.
            HtmlSaveOptions options = new HtmlSaveOptions();

            // If we save the document normally, there will be one output HTML
            // document with all the source document's contents.
            // Set the "DocumentSplitCriteria" property to "DocumentSplitCriteria.SectionBreak" to
            // save our document to multiple HTML files: one for each section.
            options.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;

            // Assign a custom callback to the "DocumentPartSavingCallback" property to alter the document part saving logic.
            options.DocumentPartSavingCallback = new SavedDocumentPartRename(outFileName, options.DocumentSplitCriteria);

            // If we convert a document that contains images into html, we will end up with one html file which links to several images.
            // Each image will be in the form of a file in the local file system.
            // There is also a callback that can customize the name and file system location of each image.
            options.ImageSavingCallback = new SavedImageRename(outFileName);

            doc.Save(ArtifactsDir + outFileName, options);
        }

        /// <summary>
        /// Sets custom filenames for output documents that the saving operation splits a document into.
        /// </summary>
        private class SavedDocumentPartRename : IDocumentPartSavingCallback
        {
            public SavedDocumentPartRename(string outFileName, DocumentSplitCriteria documentSplitCriteria)
            {
                mOutFileName = outFileName;
                mDocumentSplitCriteria = documentSplitCriteria;
            }

            void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
            {
                // We can access the entire source document via the "Document" property.
                Assert.True(args.Document.OriginalFileName.EndsWith("Rendering.docx"));

                string partType = string.Empty;

                switch (mDocumentSplitCriteria)
                {
                    case DocumentSplitCriteria.PageBreak:
                        partType = "Page";
                        break;
                    case DocumentSplitCriteria.ColumnBreak:
                        partType = "Column";
                        break;
                    case DocumentSplitCriteria.SectionBreak:
                        partType = "Section";
                        break;
                    case DocumentSplitCriteria.HeadingParagraph:
                        partType = "Paragraph from heading";
                        break;
                }

                string partFileName = $"{mOutFileName} part {++mCount}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";

                // Below are two ways of specifying where Aspose.Words will save each part of the document.
                // 1 -  Set a filename for the output part file:
                args.DocumentPartFileName = partFileName;

                // 2 -  Create a custom stream for the output part file:
                args.DocumentPartStream = new FileStream(ArtifactsDir + partFileName, FileMode.Create);

                Assert.True(args.DocumentPartStream.CanWrite);
                Assert.False(args.KeepDocumentPartStreamOpen);
            }

            private int mCount;
            private readonly string mOutFileName;
            private readonly DocumentSplitCriteria mDocumentSplitCriteria;
        }

        /// <summary>
        /// Sets custom filenames for image files that an HTML conversion creates.
        /// </summary>
        public class SavedImageRename : IImageSavingCallback
        {
            public SavedImageRename(string outFileName)
            {
                mOutFileName = outFileName;
            }

            void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
            {
                string imageFileName = $"{mOutFileName} shape {++mCount}, of type {args.CurrentShape.ShapeType}{Path.GetExtension(args.ImageFileName)}";

                // Below are two ways of specifying where Aspose.Words will save each part of the document.
                // 1 -  Set a filename for the output image file:
                args.ImageFileName = imageFileName;

                // 2 -  Create a custom stream for the output image file:
                args.ImageStream = new FileStream(ArtifactsDir + imageFileName, FileMode.Create);

                Assert.True(args.ImageStream.CanWrite);
                Assert.True(args.IsImageAvailable);
                Assert.False(args.KeepImageStreamOpen);
            }

            private int mCount;
            private readonly string mOutFileName;
        }
        //ExEnd

        //ExStart
        //ExFor:CssSavingArgs
        //ExFor:CssSavingArgs.CssStream
        //ExFor:CssSavingArgs.Document
        //ExFor:CssSavingArgs.IsExportNeeded
        //ExFor:CssSavingArgs.KeepCssStreamOpen
        //ExFor:CssStyleSheetType
        //ExFor:HtmlSaveOptions.CssSavingCallback
        //ExFor:HtmlSaveOptions.CssStyleSheetFileName
        //ExFor:HtmlSaveOptions.CssStyleSheetType
        //ExFor:ICssSavingCallback
        //ExFor:ICssSavingCallback.CssSaving(CssSavingArgs)
        //ExSummary:Shows how to work with CSS stylesheets that an HTML conversion creates.
        [Test] //ExSkip
        public void ExternalCssFilenames()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we convert the document to HTML.
            HtmlSaveOptions options = new HtmlSaveOptions();

            // Set the "CssStylesheetType" property to "CssStyleSheetType.External" to
            // accompany a saved HTML document with an external CSS stylesheet file.
            options.CssStyleSheetType = CssStyleSheetType.External;

            // Below are two ways of specifying directories and filenames for output CSS stylesheets.
            // 1 -  Use the "CssStyleSheetFileName" property to assign a filename to our stylesheet:
            options.CssStyleSheetFileName = ArtifactsDir + "SavingCallback.ExternalCssFilenames.css";

            // 2 -  Use a custom callback to name our stylesheet:
            options.CssSavingCallback =
                new CustomCssSavingCallback(ArtifactsDir + "SavingCallback.ExternalCssFilenames.css", true, false);

            doc.Save(ArtifactsDir + "SavingCallback.ExternalCssFilenames.html", options);
        }

        /// <summary>
        /// Sets a custom filename, along with other parameters for an external CSS stylesheet.
        /// </summary>
        private class CustomCssSavingCallback : ICssSavingCallback
        {
            public CustomCssSavingCallback(string cssDocFilename, bool isExportNeeded, bool keepCssStreamOpen)
            {
                mCssTextFileName = cssDocFilename;
                mIsExportNeeded = isExportNeeded;
                mKeepCssStreamOpen = keepCssStreamOpen;
            }

            public void CssSaving(CssSavingArgs args)
            {
                // We can access the entire source document via the "Document" property.
                Assert.True(args.Document.OriginalFileName.EndsWith("Rendering.docx"));

                args.CssStream = new FileStream(mCssTextFileName, FileMode.Create);
                args.IsExportNeeded = mIsExportNeeded;
                args.KeepCssStreamOpen = mKeepCssStreamOpen;

                Assert.True(args.CssStream.CanWrite);
            }

            private readonly string mCssTextFileName;
            private readonly bool mIsExportNeeded;
            private readonly bool mKeepCssStreamOpen;
        }
        //ExEnd
    }
}