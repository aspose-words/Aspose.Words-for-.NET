// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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

        [Test]
        public void PageFileNameSavingCallback()
        {
            //ExStart
            //ExFor:IPageSavingCallback
            //ExFor:PageSavingArgs
            //ExFor:PageSavingArgs.PageFileName
            //ExFor:FixedPageSaveOptions.PageSavingCallback
            //ExSummary:Shows how separate pages are saved when a document is exported to fixed page format.
            Document doc = new Document(MyDir + "Rendering.doc");

            HtmlFixedSaveOptions htmlFixedSaveOptions =
                new HtmlFixedSaveOptions { PageIndex = 0, PageCount = doc.PageCount };
            htmlFixedSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            doc.Save(ArtifactsDir + "Rendering.html", htmlFixedSaveOptions);

            string[] filePaths = Directory.GetFiles(ArtifactsDir + "", "Page_*.html");

            for (int i = 0; i < doc.PageCount; i++)
            {
                string file = string.Format(ArtifactsDir + "Page_{0}.html", i);
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
                // Specify name of the output file for the current page.
                args.PageFileName = string.Format(ArtifactsDir + "Page_{0}.html", args.PageIndex);
            }
        }
        //ExEnd

        //ExStart
        //ExFor:IDocumentPartSavingCallback
        //ExFor:IDocumentPartSavingCallback(DocumentPartSavingArgs)
        //ExFor:IImageSavingCallback
        //ExFor:IImageSavingCallback.ImageSaving
        //ExFor:ImageSavingArgs
        //ExFor:ImageSavingArgs.ImageFileName
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.ImageSavingCallback
        //ExSummary:Shows how split a document into parts and save them.
        [Test] //ExSkip
        public void DocumentParts()
        {
            // Open a document to be converted to html
            Document doc = new Document(MyDir + "Rendering.doc");
            string outFileName = "SavingCallback.DocumentParts.html";

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
        /// Renames saved document parts that are produced when an HTML document is saved while being split according to a criteria
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
                
                args.DocumentPartFileName = $"{mOutFileName} part {++mCount}, part type {partType}{Path.GetExtension(args.DocumentPartFileName)}";
            }

            private int mCount;
            private readonly string mOutFileName;
            private readonly DocumentSplitCriteria mDocumentSplitCriteria;
        }

        /// <summary>
        /// Renames saved images that are produced when an HTML document is saved 
        /// </summary>
        public class SavedImageRename : IImageSavingCallback
        {
            public SavedImageRename(string outFileName)
            {
                mOutFileName = outFileName;
            }

            void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
            {
                args.ImageFileName = $"{mOutFileName} shape {++mCount}, shape type {args.CurrentShape.ShapeType}{Path.GetExtension(args.ImageFileName)}";
            }

            private int mCount;
            private readonly string mOutFileName;
        }
        //ExEnd
    }
}