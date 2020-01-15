// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
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
            htmlFixedSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            PsSaveOptions psSaveOptions = new PsSaveOptions();
            psSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            SvgSaveOptions svgSaveOptions = new SvgSaveOptions();
            svgSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            XamlFixedSaveOptions xamlFixedSaveOptions = new XamlFixedSaveOptions();
            xamlFixedSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            XpsSaveOptions xpsSaveOptions = new XpsSaveOptions();
            xpsSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();
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
        //ExSummary:Shows how separate pages are saved when a document is exported to fixed page format.
        [Test] //ExSkip
        public void PageFileName()
        {
            Document doc = new Document(MyDir + "Rendering.doc");

            HtmlFixedSaveOptions htmlFixedSaveOptions =
                new HtmlFixedSaveOptions { PageIndex = 0, PageCount = doc.PageCount };
            htmlFixedSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            doc.Save($"{ArtifactsDir}SavingCallback.PageFileName.html", htmlFixedSaveOptions);

            string[] filePaths = Directory.GetFiles(ArtifactsDir, "SavingCallback.PageFileName.Page_*.html");

            for (int i = 0; i < doc.PageCount; i++)
            {
                string file = $"{ArtifactsDir}SavingCallback.PageFileName.Page_{i}.html";
                Assert.AreEqual(file, filePaths[i]);//ExSkip
            }
        }

        /// <summary>
        /// Custom PageFileName is specified.
        /// </summary>
        private class CustomPageFileNamePageSavingCallback : IPageSavingCallback
        {
            public void PageSaving(PageSavingArgs args)
            {
                string outFileName = $"{ArtifactsDir}SavingCallback.PageFileName.Page_{args.PageIndex}.html";

                // Specify name of the output file for the current page either in this 
                args.PageFileName = outFileName;

                // ..or by setting up a custom stream
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
        //ExSummary:Shows how split a document into parts and save them.
        [Test] //ExSkip
        public void DocumentParts()
        {
            // Open a document to be converted to html
            Document doc = new Document(MyDir + "Rendering.doc");
            string outFileName = "SavingCallback.DocumentParts.Rendering.html";

            // We can use an appropriate SaveOptions subclass to customize the conversion process
            HtmlSaveOptions options = new HtmlSaveOptions();

            // We can use it to split a document into smaller parts, in this instance split by section breaks
            // Each part will be saved into a separate file, creating many files during the conversion process instead of just one
            options.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;

            // We can set a callback to name each document part file ourselves
            options.DocumentPartSavingCallback = new SavedDocumentPartRename(outFileName, options.DocumentSplitCriteria);

            // If we convert a document that contains images into html, we will end up with one html file which links to several images
            // Each image will be in the form of a file in the local file system
            // There is also a callback that can customize the name and file system location of each image
            options.ImageSavingCallback = new SavedImageRename(outFileName);

            // The DocumentPartSaving() and ImageSaving() methods of our callbacks will be run at this time
            doc.Save(ArtifactsDir + outFileName, options);
        }

        /// <summary>
        /// Renames saved document parts that are produced when an HTML document is saved while being split according to a criteria.
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
                Assert.True(args.Document.OriginalFileName.EndsWith("Rendering.doc"));

                string partType = "";

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

                // We can designate the filename and location of each output file either by filename
                args.DocumentPartFileName = partFileName;

                // Or we can make a new stream and choose the location of the file at construction
                args.DocumentPartStream = new FileStream(ArtifactsDir + partFileName, FileMode.Create);
                Assert.True(args.DocumentPartStream.CanWrite);
                Assert.False(args.KeepDocumentPartStreamOpen);
            }

            private int mCount;
            private readonly string mOutFileName;
            private readonly DocumentSplitCriteria mDocumentSplitCriteria;
        }

        /// <summary>
        /// Renames saved images that are produced when an HTML document is saved.
        /// </summary>
        public class SavedImageRename : IImageSavingCallback
        {
            public SavedImageRename(string outFileName)
            {
                mOutFileName = outFileName;
            }

            void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
            {
                // Same filename and stream functions as above in IDocumentPartSavingCallback apply here
                string imageFileName = $"{mOutFileName} shape {++mCount}, of type {args.CurrentShape.ShapeType}{Path.GetExtension(args.ImageFileName)}";

                args.ImageFileName = imageFileName;

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
        //ExSummary:Shows how to work with CSS stylesheets that may be created along with Html documents.
        [Test] //ExSkip
        public void CssSavingCallback()
        {
            // Open a document to be converted to html
            Document doc = new Document(MyDir + "Rendering.doc");

            // If our output document will produce a CSS stylesheet, we can use an HtmlSaveOptions to control where it is saved
            HtmlSaveOptions options = new HtmlSaveOptions();

            // By default, a CSS stylesheet is stored inside its HTML document, but we can have it saved to a separate file
            options.CssStyleSheetType = CssStyleSheetType.External;

            // We can designate a filename for our stylesheet like this
            options.CssStyleSheetFileName = ArtifactsDir + "SavingCallback.CssSavingCallback.css";

            // A custom ICssSavingCallback implementation can also control where that stylesheet will be saved and linked to by the Html document
            // This callback will override the filename we specified above in options.CssStyleSheetFileName,
            // but will give us more control over the saving process
            options.CssSavingCallback =
                new CustomCssSavingCallback(ArtifactsDir + "SavingCallback.CssSavingCallback.css", true, false);

            // The CssSaving() method of our callback will be called at this stage
            doc.Save(ArtifactsDir + "SavingCallback.CssSavingCallback.html", options);
        }

        /// <summary>
        /// Designates a filename and other parameters for the saving of a CSS stylesheet
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
                // Set up the stream that will create the CSS document         
                args.CssStream = new FileStream(mCssTextFileName, FileMode.Create);
                Assert.True(args.CssStream.CanWrite);
                args.IsExportNeeded = mIsExportNeeded;
                args.KeepCssStreamOpen = mKeepCssStreamOpen;

                // We can also access the original document here like this
                Assert.True(args.Document.OriginalFileName.EndsWith("Rendering.doc"));
            }

            private readonly string mCssTextFileName;
            private readonly bool mIsExportNeeded;
            private readonly bool mKeepCssStreamOpen;
        }
        //ExEnd
    }
}