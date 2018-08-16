// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using NUnit.Framework;
using Aspose.Words.Saving;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        #region PageMargins

        //For assert this test you need to open HTML docs and they shouldn't have negative left margins
        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportPageMargins(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.SaveFormat = saveFormat;
            saveOptions.ExportPageMargins = true;

            Save(doc, @"\Artifacts\HtmlSaveOptions.ExportPageMargins." + saveFormat.ToString().ToLower(), saveFormat, saveOptions);
        }

        #endregion

        #region HtmlOfficeMathOutputMode

        [Test]
        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.Image)]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.MathML)]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.Text)]
        public void ExportOfficeMath(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
        {
            Document doc = new Document(MyDir + "OfficeMath.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.OfficeMathOutputMode = outputMode;

            Save(doc, @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(), saveFormat, saveOptions);

            switch (saveFormat)
            {
                case SaveFormat.Html:
                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(), "<img src=\"HtmlSaveOptions.ExportToHtmlUsingImage.001.png\" width=\"49\" height=\"19\" alt=\"\" style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />");
                    return;

                case SaveFormat.Mhtml:
                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(), "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>A</mi><mo>=</mo><mi>π</mi><msup><mrow><mi>r</mi></mrow><mrow><mn>2</mn></mrow></msup></math>");
                    return;

                case SaveFormat.Epub:
                    DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(), "<span style=\"font-family:\'Cambria Math\'\">A=π</span><span style=\"font-family:\'Cambria Math\'\">r</span><span style=\"font-family:\'Cambria Math\'\">2</span>");
                    return;
            }
        }

        #endregion

        private static void Save(Document inputDoc, string outputDocPath, SaveFormat saveFormat,
            SaveOptions saveOptions)
        {
            switch (saveFormat)
            {
                case SaveFormat.Html:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions);
                    return;
                case SaveFormat.Mhtml:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions);
                    return;
                case SaveFormat.Epub:
                    inputDoc.Save(MyDir + outputDocPath, saveOptions);
                    return;
            }

            return;
        }

        [Test]
        public void ControlListLabelsExportToHtml()
        {
            Document doc = new Document(MyDir + "Lists.PrintOutAllLists.doc");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // This option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
            // otherwise HTML <p> tag is used. This is also the default value.
            saveOptions.ExportListLabels = ExportListLabels.Auto;
            doc.Save(MyDir + @"\Artifacts\Document.ExportListLabels Auto.html", saveOptions);

            // Using this option the <p> tag is used for any list label representation.
            saveOptions.ExportListLabels = ExportListLabels.AsInlineText;
            doc.Save(MyDir + @"\Artifacts\Document.ExportListLabels InlineText.html", saveOptions);

            // The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
            saveOptions.ExportListLabels = ExportListLabels.ByHtmlTags;
            doc.Save(MyDir + @"\Artifacts\Document.ExportListLabels HtmlTags.html", saveOptions);
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportUrlForLinkedImage(bool export)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportUrlForLinkedImage.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportOriginalUrlForLinkedImages = export;

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

            string[] dirFiles = Directory.GetFiles(MyDir + @"\Artifacts\", "HtmlSaveOptions.ExportUrlForLinkedImage.001.png", SearchOption.AllDirectories);

            if (dirFiles.Length == 0)
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
            else
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
        }

        [Ignore("Bug, css styles starting with -aw, even if ExportRoundtripInformation is false")]
        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportRoundtripInformation(bool valueHtml)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportRoundtripInformation = valueHtml;

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html");

            if (valueHtml)
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html", "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
            else
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html", "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
        }

        [Test]
        public void RoundtripInformationDefaulValue()
        {
            //Assert that default value is true for HTML and false for MHTML and EPUB.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            Assert.AreEqual(true, saveOptions.ExportRoundtripInformation);

            saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            Assert.AreEqual(false, saveOptions.ExportRoundtripInformation);

            saveOptions = new HtmlSaveOptions(SaveFormat.Epub);
            Assert.AreEqual(false, saveOptions.ExportRoundtripInformation);
        }

        [Test]
        public void ConfigForSavingExternalResources()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.ResourceFolder = "Resources";
            saveOptions.ResourceFolderAlias = "https://www.aspose.com/";

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ExportPageMargins Out.html", saveOptions);

            string[] imageFiles = Directory.GetFiles(MyDir + @"\Artifacts\Resources\", "*.png", SearchOption.AllDirectories);
            Assert.AreEqual(3, imageFiles.Length);

            string[] fontFiles = Directory.GetFiles(MyDir + @"\Artifacts\Resources\", "*.ttf", SearchOption.AllDirectories);
            Assert.AreEqual(1, fontFiles.Length);

            string[] cssFiles = Directory.GetFiles(MyDir + @"\Artifacts\Resources\", "*.css", SearchOption.AllDirectories);
            Assert.AreEqual(1, cssFiles.Length);

            DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportPageMargins Out.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExportPageMargins Out.css\"");
        }

        [Test]
        public void ConvertFontsAsBase64()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ResourceFolder = "Resources";
            saveOptions.ExportFontResources = true;
            saveOptions.ExportFontsAsBase64 = true;
            
            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ExportPageMargins Out.html", saveOptions);
}
        [TestCase(HtmlVersion.Html5)]
        [TestCase(HtmlVersion.Xhtml)]
        public void Html5Support(HtmlVersion htmlVersion)
        {
            Document doc = new Document(MyDir + "Document.doc");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.HtmlVersion = htmlVersion;
        }

        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void ExportFonts(bool exportAsBase64)
        {
            Document doc = new Document(MyDir + "Document.doc");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportFontResources = true;
            saveOptions.ExportFontsAsBase64 = exportAsBase64;

            switch (exportAsBase64)
            {
                case false:

                    doc.Save(MyDir + @"\Artifacts\DocumentExportFonts Out 1.html", saveOptions);
                    Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\", "DocumentExportFonts Out 1.times.ttf", SearchOption.AllDirectories)); //Verify that the font has been added to the folder
                    break;

                case true:

                    doc.Save(MyDir + @"\Artifacts\DocumentExportFonts Out 2.html", saveOptions);
                    Assert.IsEmpty(Directory.GetFiles(MyDir + @"\Artifacts\", "DocumentExportFonts Out 2.times.ttf", SearchOption.AllDirectories)); //Verify that the font is not added to the folder
                    break;
            }
        }

        [Test]
        public void ResourceFolderPriority()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ResourceFolder.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.ResourceFolder = MyDir + @"\Artifacts\Resources";
            saveOptions.ResourceFolderAlias = "http://example.com/resources";

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ResourceFolder Out.html", saveOptions);

            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources", "HtmlSaveOptions.ResourceFolder Out.001.jpeg", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources", "HtmlSaveOptions.ResourceFolder Out.002.png", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources", "HtmlSaveOptions.ResourceFolder Out.calibri.ttf", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources", "HtmlSaveOptions.ResourceFolder Out.css", SearchOption.AllDirectories));

        }

        [Test]
        public void ResourceFolderLowPriority()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ResourceFolder.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.FontsFolder = MyDir + @"\Artifacts\Fonts";
            saveOptions.ImagesFolder = MyDir + @"\Artifacts\Images";
            saveOptions.ResourceFolder = MyDir + @"\Artifacts\Resources";
            saveOptions.ResourceFolderAlias = "http://example.com/resources";

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ResourceFolder Out.html", saveOptions);

            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Images", "HtmlSaveOptions.ResourceFolder Out.001.jpeg", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Images", "HtmlSaveOptions.ResourceFolder Out.002.png", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Fonts", "HtmlSaveOptions.ResourceFolder Out.calibri.ttf", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources", "HtmlSaveOptions.ResourceFolder Out.css", SearchOption.AllDirectories));
        }
    }
}
