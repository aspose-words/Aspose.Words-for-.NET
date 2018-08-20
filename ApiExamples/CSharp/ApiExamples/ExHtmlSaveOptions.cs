// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using NUnit.Framework;
using Aspose.Words.Saving;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        // Note: For assert this test you need to open HTML docs and they shouldn't have negative left margins
        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportPageMargins(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                SaveFormat = saveFormat,
                ExportPageMargins = true
            };

            Save(doc, @"\Artifacts\HtmlSaveOptions.ExportPageMargins." + saveFormat.ToString().ToLower(), saveFormat,
                saveOptions);
        }

        [Test]
        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.Image)]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.MathML)]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.Text)]
        public void ExportOfficeMath(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
        {
            Document doc = new Document(MyDir + "OfficeMath.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                OfficeMathOutputMode = outputMode
            };

            Save(doc, @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(),
                saveFormat, saveOptions);

            switch (saveFormat)
            {
                case SaveFormat.Html:
                    DocumentHelper.FindTextInFile(
                        MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(),
                        "<img src=\"HtmlSaveOptions.ExportToHtmlUsingImage.001.png\" width=\"49\" height=\"19\" alt=\"\" style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />");
                    return;

                case SaveFormat.Mhtml:
                    DocumentHelper.FindTextInFile(
                        MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(),
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>A</mi><mo>=</mo><mi>π</mi><msup><mrow><mi>r</mi></mrow><mrow><mn>2</mn></mrow></msup></math>");
                    return;

                case SaveFormat.Epub:
                    DocumentHelper.FindTextInFile(
                        MyDir + @"\Artifacts\HtmlSaveOptions.ExportToHtmlUsingImage." + saveFormat.ToString().ToLower(),
                        "<span style=\"font-family:\'Cambria Math\'\">A=π</span><span style=\"font-family:\'Cambria Math\'\">r</span><span style=\"font-family:\'Cambria Math\'\">2</span>");
                    return;
            }
        }

        [Test]
        [TestCase(ExportListLabels.Auto)]
        [TestCase(ExportListLabels.AsInlineText)]
        [TestCase(ExportListLabels.ByHtmlTags)]
        public void ControlListLabelsExportToHtml(ExportListLabels howExportListLabels)
        {
            Document doc = new Document(MyDir + "Lists.PrintOutAllLists.doc");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
                // otherwise HTML <p> tag is used. This is also the default value.
                // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation.
                // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
                ExportListLabels = howExportListLabels
            };

            doc.Save(MyDir + @"\Artifacts\Document.ExportListLabels.html", saveOptions);
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportUrlForLinkedImage(bool export)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportUrlForLinkedImage.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportOriginalUrlForLinkedImages = export };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

            string[] dirFiles = Directory.GetFiles(MyDir + @"\Artifacts\",
                "HtmlSaveOptions.ExportUrlForLinkedImage.001.png", SearchOption.AllDirectories);

            if (dirFiles.Length == 0)
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html",
                    "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
            else
                DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportUrlForLinkedImage.html",
                    "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
        }

        [Test]
        public void ExportRoundtripInformation()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportRoundtripInformation = true };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.RoundtripInformation.html", saveOptions);
        }

        [Test]
        public void ConfigForSavingExternalResources()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                ResourceFolder = "Resources",
                ResourceFolderAlias = "https://www.aspose.com/"
            };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ExportPageMargins.html", saveOptions);

            string[] imageFiles =
                Directory.GetFiles(MyDir + @"\Artifacts\Resources\", "*.png", SearchOption.AllDirectories);
            Assert.AreEqual(3, imageFiles.Length);

            string[] fontFiles =
                Directory.GetFiles(MyDir + @"\Artifacts\Resources\", "*.ttf", SearchOption.AllDirectories);
            Assert.AreEqual(1, fontFiles.Length);

            string[] cssFiles =
                Directory.GetFiles(MyDir + @"\Artifacts\Resources\", "*.css", SearchOption.AllDirectories);
            Assert.AreEqual(1, cssFiles.Length);

            DocumentHelper.FindTextInFile(MyDir + @"\Artifacts\HtmlSaveOptions.ExportPageMargins.html",
                "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExportPageMargins.css\"");
        }

        [Test]
        public void ConvertFontsAsBase64()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ResourceFolder = "Resources",
                ExportFontResources = true,
                ExportFontsAsBase64 = true
            };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ExportPageMargins.html", saveOptions);
        }

        [TestCase(HtmlVersion.Html5)]
        [TestCase(HtmlVersion.Xhtml)]
        public void Html5Support(HtmlVersion htmlVersion)
        {
            Document doc = new Document(MyDir + "Document.doc");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                HtmlVersion = htmlVersion
            };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.Html5Support.html", saveOptions);
        }

        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void ExportFonts(bool exportAsBase64)
        {
            Document doc = new Document(MyDir + "Document.doc");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,
                ExportFontsAsBase64 = exportAsBase64
            };

            switch (exportAsBase64)
            {
                case false:

                    doc.Save(MyDir + @"\Artifacts\DocumentExportFonts.html", saveOptions);
                    // Verify that the font has been added to the folder
                    Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"Artifacts\", "DocumentExportFonts.times.ttf",
                        SearchOption.AllDirectories));
                    break;

                case true:

                    doc.Save(MyDir + @"\Artifacts\DocumentExportFonts.html", saveOptions);
                    // Verify that the font is not added to the folder
                    Assert.IsEmpty(Directory.GetFiles(MyDir + @"Artifacts\", "DocumentExportFonts.times.ttf",
                        SearchOption.AllDirectories));
                    break;
            }
        }

        [Test]
        public void ResourceFolderPriority()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ResourceFolder.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                ResourceFolder = MyDir + @"\Artifacts\Resources",
                ResourceFolderAlias = "http://example.com/resources"
            };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ResourceFolder.html", saveOptions);

            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources",
                "HtmlSaveOptions.ResourceFolder.001.jpeg", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources",
                "HtmlSaveOptions.ResourceFolder.002.png", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources",
                "HtmlSaveOptions.ResourceFolder.calibri.ttf", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources", "HtmlSaveOptions.ResourceFolder.css",
                SearchOption.AllDirectories));
        }

        [Test]
        public void ResourceFolderLowPriority()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ResourceFolder.docx");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                FontsFolder = MyDir + @"\Artifacts\Fonts",
                ImagesFolder = MyDir + @"\Artifacts\Images",
                ResourceFolder = MyDir + @"\Artifacts\Resources",
                ResourceFolderAlias = "http://example.com/resources"
            };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.ResourceFolder.html", saveOptions);

            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Images",
                "HtmlSaveOptions.ResourceFolder.001.jpeg", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Images", "HtmlSaveOptions.ResourceFolder.002.png",
                SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Fonts",
                "HtmlSaveOptions.ResourceFolder.calibri.ttf", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(MyDir + @"\Artifacts\Resources", "HtmlSaveOptions.ResourceFolder.css",
                SearchOption.AllDirectories));
        }

        [Test]
        public void SvgMetafileFormat()
        {
            DocumentBuilder builder = new DocumentBuilder();

            builder.Write("Here is an SVG image: ");
            builder.InsertHtml(
                @"<svg height='210' width='500'>
                    <polygon points='100,10 40,198 190,78 10,78 160,198' 
                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
                  </svg> ");

            builder.Document.Save(MyDir + @"\Artifacts\HtmlSaveOptions.MetafileFormat.html",
                new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.Png });
        }

        [Test]
        public void PngMetafileFormat()
        {
            DocumentBuilder builder = new DocumentBuilder();

            builder.Write("Here is an Png image: ");
            builder.InsertHtml(
                @"<svg height='210' width='500'>
                    <polygon points='100,10 40,198 190,78 10,78 160,198' 
                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
                  </svg> ");

            builder.Document.Save(MyDir + @"\Artifacts\HtmlSaveOptions.MetafileFormat.html",
                new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.Png });
        }

        [Test]
        public void EmfOrWmfMetafileFormat()
        {
            DocumentBuilder builder = new DocumentBuilder();

            builder.Write("Here is an image as is: ");
            builder.InsertHtml(
                @"<img src=""data:image/png;base64,
                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP
                    C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA
                    AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J
                    REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq
                    ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0
                    vr4MkhoXe0rZigAAAABJRU5ErkJggg=="" alt=""Red dot"" />");

            builder.Document.Save(MyDir + @"\Artifacts\HtmlSaveOptions.MetafileFormat.html",
                new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.EmfOrWmf });
        }

        [Test]
        public void CssClassNamesPrefix()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.CssClassNamePrefix
            //ExSummary: Shows how to specifies a prefix which is added to all CSS class names
            Document doc = new Document(MyDir + "HtmlSaveOptions.CssClassNamePrefix.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.Embedded,
                CssClassNamePrefix = "aspose-"
            };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
            //ExEnd
        }

        [Test]
        public void CssClassNamesNotValidPrefix()
        {
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            Assert.Throws<ArgumentException>(() => saveOptions.CssClassNamePrefix = "@%-",
                "The class name prefix must be a valid CSS identifier.");
        }

        [Test]
        public void CssClassNamesNullPrefix()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.CssClassNamePrefix.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.Embedded,
                CssClassNamePrefix = null
            };

            doc.Save(MyDir + @"\Artifacts\HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
        }

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
        }
    }
}