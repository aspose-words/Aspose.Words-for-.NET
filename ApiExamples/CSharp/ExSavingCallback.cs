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

            SwfSaveOptions swfSaveOptions = new SwfSaveOptions();
            swfSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

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
            //ExFor:HtmlFixedSaveOptions.PageSavingCallback
            //ExSummary:Shows how separate pages are saved when a document is exported to fixed page format.
            Document doc = new Document(MyDir + "Rendering.doc");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions { PageIndex = 0, PageCount = doc.PageCount };
            htmlFixedSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            doc.Save(MyDir + @"\Artifacts\Rendering.html", htmlFixedSaveOptions);

            string[] filePaths = Directory.GetFiles(MyDir + @"\Artifacts\", "Page_*.html");

            for (int i = 0; i < doc.PageCount; i++)
            {
                string file = string.Format(MyDir + @"\Artifacts\Page_{0}.html", i);
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
                args.PageFileName = string.Format(MyDir + @"\Artifacts\Page_{0}.html", args.PageIndex);
            }
        }
        //ExEnd
    }
}