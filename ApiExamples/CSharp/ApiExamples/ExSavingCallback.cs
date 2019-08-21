using System.IO;
using System.Threading;
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
        //ExFor:PageSavingArgs
        //ExFor:PageSavingArgs.PageFileName
        //ExFor:FixedPageSaveOptions.PageSavingCallback
        //ExSummary:Shows how separate pages are saved when a document is exported to fixed page format.
        [Test] //ExSkip
        public void PageFileNameSavingCallback()
        {
            Document doc = new Document(MyDir + "Rendering.doc");

            HtmlFixedSaveOptions htmlFixedSaveOptions =
                new HtmlFixedSaveOptions { PageIndex = 0, PageCount = doc.PageCount };
            htmlFixedSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            doc.Save(ArtifactsDir + "Rendering.html", htmlFixedSaveOptions);

            string[] filePaths = Directory.GetFiles(ArtifactsDir + "", "Page_*.html");

            for (int i = 0; i < doc.PageCount; i++)
            {
                string file = string.Format(ArtifactsDir + "Page_{0}.html", i);
                Assert.AreEqual(file, filePaths[i]); //ExSkip
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
        //ExFor:CssSavingArgs
        //ExFor:CssSavingArgs.CssStream
        //ExFor:CssSavingArgs.Document
        //ExFor:CssSavingArgs.IsExportNeeded
        //ExFor:CssSavingArgs.KeepCssStreamOpen
        //ExFor:CssStyleSheetType
        //ExFor:ICssSavingCallback
        //ExFor:ICssSavingCallback.CssSaving(CssSavingArgs)
        //ExSummary:Shows how to work with CSS stylesheets that may be created along with Html documents.
        [Test] //ExSkip
        public void CssSavingCallback()
        {
            // Open a document to be converted to html
            Document doc = new Document(MyDir + "Rendering.doc");

            // If our output document will produce a CSS stylesheet, we can use an HtmlSaveOptions to control where it is saved
            HtmlSaveOptions htmlFixedSaveOptions = new HtmlSaveOptions();

            // By default, a CSS stylesheet are stored inside its HTML document, but we can have it saved to a separate file
            htmlFixedSaveOptions.CssStyleSheetType = CssStyleSheetType.External;

            // A custom ICssSavingCallback implementation can control where that stylesheet will be saved and linked to by the Html document
            htmlFixedSaveOptions.CssSavingCallback =
                new CustomCssSavingCallback(ArtifactsDir + "Rendering.CssSavingCallback.css", true, false);

            // The CssSaving() method of our callback will be called at this stage
            doc.Save(ArtifactsDir + "Rendering.CssSavingCallback.html", htmlFixedSaveOptions);
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