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
            Document doc = new Document(MyDir + "Rendering.doc");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions { PageIndex = 0, PageCount = doc.PageCount };
            htmlFixedSaveOptions.PageSavingCallback = new CustomPageFileNamePageSavingCallback();

            doc.Save(MyDir + @"\Artifacts\out.html", htmlFixedSaveOptions);

            string[] filePaths = Directory.GetFiles(MyDir, "Page_*.html");

            for (int i = 0; i < doc.PageCount; i++)
            {
                string file = string.Format(MyDir + "Page_{0}.html", i);
                Assert.AreEqual(file, filePaths[i]);
            }
        }

        [Test]
        public void PageStreamSavingCallback()
        {
            Stream docStream = new FileStream(MyDir + "Rendering.doc", FileMode.Open);
            Document doc = new Document(docStream);

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions { PageIndex = 0, PageCount = doc.PageCount };
            htmlFixedSaveOptions.PageSavingCallback = new CustomPageStreamPageSavingCallback();

            doc.Save(MyDir + @"\Artifacts\out.html", htmlFixedSaveOptions);

            docStream.Close();
        }

        /// <summary>
        /// Custom PageFileName is specified.
        /// </summary>
        private class CustomPageFileNamePageSavingCallback : IPageSavingCallback
        {
            public void PageSaving(PageSavingArgs args)
            {
                // Specify name of the output file for the current page.
                args.PageFileName = string.Format(MyDir + "Page_{0}.html", args.PageIndex);
            }
        }

        /// <summary>
        /// Custom PageStream is specified.
        /// </summary>
        private class CustomPageStreamPageSavingCallback : IPageSavingCallback
        {
            public void PageSaving(PageSavingArgs args)
            {
                // Specify memory stream for the current page.
                args.PageStream = new MemoryStream();
                args.KeepPageStreamOpen = true;
            }
        }
    }
}
