﻿using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace QA_Tests.Tests.SaveOptions.Html
{
    /// <summary>
    /// Tests that verify saving to htmlfixed using encoding parameter in "HtmlFixedSaveOptions"
    /// </summary>
    [TestFixture]
    internal class HtmlFixedSaveOptionsEncoding : QaTestsBase
    {
        //Note: Tests doesn't containt validation result, because it's may take a lot of time for assert result
        //For validation result, you can save the document to html file and check out with notepad++, that file encoding will be correctly displayed (Encoding tab in Notepad++)
        [Test]
        public void EncodingUsingSystemTextEncoding()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = Encoding.ASCII,
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedCss = true,
                ExportEmbeddedFonts = true,
                ExportEmbeddedImages = true,
                ExportEmbeddedSvg = true
            };

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, htmlFixedSaveOptions);

            dstStream.Dispose();
        }

        [Test]
        public void EncodingUsingNewEncoding()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = new UTF32Encoding(),
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedCss = true,
                ExportEmbeddedFonts = true,
                ExportEmbeddedImages = true,
                ExportEmbeddedSvg = true
            };

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, htmlFixedSaveOptions);

            dstStream.Dispose();
        }


        [Test]
        public void EncodingUsingGetEncoding()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = Encoding.GetEncoding("utf-16"),
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedCss = true,
                ExportEmbeddedFonts = true,
                ExportEmbeddedImages = true,
                ExportEmbeddedSvg = true
            };

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, htmlFixedSaveOptions);

            dstStream.Dispose();
        }
    }
}
