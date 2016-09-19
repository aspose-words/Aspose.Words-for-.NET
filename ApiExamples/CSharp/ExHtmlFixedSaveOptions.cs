// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Text;
using System;

using Aspose.Words;
using Aspose.Words.Saving;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlFixedSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseEncoding()
        {
            //ExStart
            //ExFor:Saving.HtmlFixedSaveOptions.Encoding
            //ExSummary:Shows how to use "Encoding" parameter with "HtmlFixedSaveOptions"
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello World!");

            //Create "HtmlFixedSaveOptions" with "Encoding" parameter
            //You can also set "Encoding" using System.Text.Encoding, like "Encoding.ASCII", or "Encoding.GetEncoding()"
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = new ASCIIEncoding(),
                SaveFormat = SaveFormat.HtmlFixed,
            };

            //Uses "HtmlFixedSaveOptions"
            doc.Save(MyDir + @"\Artifacts\UseEncoding.html", htmlFixedSaveOptions);
            //ExEnd
        }

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

            doc.Save(MyDir + "EncodingUsingSystemTextEncoding.html", htmlFixedSaveOptions);
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

            doc.Save(MyDir + "EncodingUsingNewEncoding.html", htmlFixedSaveOptions);
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

            doc.Save(MyDir + "EncodingUsingGetEncoding.html", htmlFixedSaveOptions);
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportFormFields(bool exportFormFields)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCheckBox("CheckBox", false, 15);

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedCss = true,
                ExportEmbeddedFonts = true,
                ExportEmbeddedImages = true,
                ExportEmbeddedSvg = true,
                ExportFormFields = exportFormFields
            };

            //For assert test result you need to open documents and check that checkbox are clickable in "ExportFormFiels.html" file and are not clickable in "WithoutExportFormFiels.html" file
            if (exportFormFields == true)
            {
                doc.Save(MyDir + "ExportFormFiels.html", htmlFixedSaveOptions);
            }
            else
            {
                doc.Save(MyDir + "WithoutExportFormFiels.html", htmlFixedSaveOptions);
            }
        }

        [Test]
        [TestCase("aw")]
        [TestCase("")]
        public void CssPrefix(string cssprefix)
        {
            Document doc = new Document(MyDir + "Bookmark.doc");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
            saveOptions.CssClassNamesPrefix = cssprefix;

            doc.Save(MyDir + @"\Artifacts\cssPrefix_Out.html", saveOptions);

            DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\cssPrefix_Out\styles.css", "div");
        }

        [Test]
        [TestCase(HtmlFixedPageHorizontalAlignment.Center)]
        [TestCase(HtmlFixedPageHorizontalAlignment.Left)]
        [TestCase(HtmlFixedPageHorizontalAlignment.Right)]
        public void HorizontalAlignment(HtmlFixedPageHorizontalAlignment horizontalAlignment)
        {
            Document doc = new Document(MyDir + "Bookmark.doc");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
            saveOptions.PageHorizontalAlignment = horizontalAlignment;

            doc.Save(MyDir + @"\Artifacts\HtmlFixedPageHorizontalAlignment.html", saveOptions);
        }

        [Test]
        public void PageMarginsException()
        {
            Document doc = new Document(MyDir + "Bookmark.doc");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
            Assert.That(() => saveOptions.PageMargins = -1, Throws.TypeOf<ArgumentException>());

            doc.Save(MyDir + @"\Artifacts\HtmlFixedPageMargins.html", saveOptions);
        }

        [TestCase(0)]
        [TestCase(10)]
        public void PageMargins(int margin)
        {
            Document doc = new Document(MyDir + "Bookmark.doc");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
            saveOptions.PageMargins = margin;

            doc.Save(MyDir + @"\Artifacts\HtmlFixedPageMargins.html", saveOptions);
        }
    }
}