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
        //ExSummary:Shows how split a document into parts and save them.
        [Test] //ExSkip
        public void DocumentPartSavingCallback()
        {
            Document doc = new Document(MyDir + "Rendering.doc");

            string outFileName = "SavingCallback.DocumentPartSavingCallback.html";

            HtmlSaveOptions options = new HtmlSaveOptions();
            options.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;
            options.DocumentPartSavingCallback = new DocumentPartRename(outFileName, options.DocumentSplitCriteria);
            doc.Save(ArtifactsDir + outFileName, options);
        }

        /// <summary>
        /// Renames saved document parts that are produced when an HTML document is saved 
        /// </summary>
        private class DocumentPartRename : IDocumentPartSavingCallback
        {
            public DocumentPartRename(string outFileName, DocumentSplitCriteria documentSplitCriteria)
            {
                mOutFileName = outFileName;
                mDocumentSplitCriteria = documentSplitCriteria;
            }

            void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
            {
                args.DocumentPartFileName = $"{mOutFileName} {mDocumentSplitCriteria} #{mCount++}{Path.GetExtension(args.DocumentPartFileName)}";
            }

            private int mCount;
            private readonly string mOutFileName;
            private readonly DocumentSplitCriteria mDocumentSplitCriteria;
        }
        //ExEnd
    }
}