// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;

using NUnit.Framework;

using System;
using System.IO;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        //For assert this test you need to open html docs and they shouldn't have negative left margins
        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportPageMargins(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            Aspose.Words.Saving.HtmlSaveOptions htmlSaveOptions = new Aspose.Words.Saving.HtmlSaveOptions
            {
                SaveFormat = saveFormat, 
                ExportPageMargins = true
            };

            switch (saveFormat)
            {
                case SaveFormat.Html:
                    doc.Save(MyDir + "ExportPageMargins.html", htmlSaveOptions);
                    break;
                case SaveFormat.Mhtml:
                    doc.Save(MyDir + "ExportPageMargins.Mhtml", htmlSaveOptions);
                    break;
                case SaveFormat.Epub:
                    doc.Save(MyDir + "ExportPageMargins.Epub", htmlSaveOptions); //There is draw images bug with epub. Need write to NSezganov
                    break;
            }
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportUrlForLinkedImage(bool export)
        {
            Document doc = new Document(MyDir + "ExportUrlForLinkedImage.docx");

            Aspose.Words.Saving.HtmlSaveOptions saveOptions = new Aspose.Words.Saving.HtmlSaveOptions();
            saveOptions.ExportOriginalUrlForLinkedImages = export;

            doc.Save(MyDir + @"\Artifacts\ExportUrlForLinkedImage.html", saveOptions);

            String[] dirFiles = Directory.GetFiles(MyDir + @"\Artifacts\", "ExportUrlForLinkedImage.001.png", SearchOption.AllDirectories);

            if (dirFiles.Length == 0)
            {
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
            }
            else
            {
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\ExportUrlForLinkedImage.html", "<img src=\"ExportUrlForLinkedImage.001.png\"");
            }
        }

        [Ignore("Need to rework, for best gold asserts")]
        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportRoundtripInformation(bool valueHtml)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            Aspose.Words.Saving.HtmlSaveOptions saveOptions = new Aspose.Words.Saving.HtmlSaveOptions();
            saveOptions.ExportRoundtripInformation = valueHtml;

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html");

            if (valueHtml)
            {
                this.CompareFiles(
                    MyDir + @"\Golds\HtmlSaveOptions.WithRoundtripInformation.html",
                    MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html");
            }
            else
            {
                this.CompareFiles(
                    MyDir + @"\Golds\HtmlSaveOptions.WithoutRoundtripInformation.html",
                    MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html");
            }
        }

        [Test]
        public void RoundtripInformationDefaulValue()
        {
            //Assert that default value is true for HTML and false for MHTML and EPUB.
            Aspose.Words.Saving.HtmlSaveOptions saveOptions = new Aspose.Words.Saving.HtmlSaveOptions(SaveFormat.Html);
            Assert.AreEqual(true, saveOptions.ExportRoundtripInformation);

            saveOptions = new Aspose.Words.Saving.HtmlSaveOptions(SaveFormat.Mhtml);
            Assert.AreEqual(false, saveOptions.ExportRoundtripInformation);

            saveOptions = new Aspose.Words.Saving.HtmlSaveOptions(SaveFormat.Epub);
            Assert.AreEqual(false, saveOptions.ExportRoundtripInformation);
        }

        private void CompareFiles(string firstPath, string secondPath)
        {
            String[] linesA = File.ReadAllLines(firstPath);
            String[] linesB = File.ReadAllLines(secondPath);

            Assert.AreEqual(linesA, linesB);
        }
    }
}
