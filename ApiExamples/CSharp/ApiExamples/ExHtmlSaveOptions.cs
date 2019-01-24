// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using NUnit.Framework;
using Aspose.Words.Saving;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
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

            doc.Save(ArtifactsDir +"HtmlSaveOptions.ExportPageMargins" + FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);
        }

        [Test]
        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.Image, Category = "SkipMono")]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.MathML, Category = "SkipMono")]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.Text, Category = "SkipMono")]
        public void ExportOfficeMath(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
        {
            Document doc = new Document(MyDir + "OfficeMath.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.OfficeMathOutputMode = outputMode;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportToHtmlUsingImage" + FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);
        }

        [Test]
        [TestCase(SaveFormat.Html, true, Description = "TextBox as svg (html)")]
        [TestCase(SaveFormat.Epub, true, Description = "TextBox as svg (epub)")]
        [TestCase(SaveFormat.Mhtml, false, Description = "TextBox as img (mhtml)")]
        public void ExportTextBoxAsSvg(SaveFormat saveFormat, bool isTextBoxAsSvg)
        {
            string[] dirFiles;

            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportTextBoxAsSvg.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(saveFormat);
            saveOptions.ExportTextBoxAsSvg = isTextBoxAsSvg;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTextBoxAsSvg" + FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);

            switch (saveFormat)
            {
                case SaveFormat.Html:
                    
                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.AllDirectories);
                    Assert.IsEmpty(dirFiles);
                    return;

                case SaveFormat.Epub:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.AllDirectories);
                    Assert.IsEmpty(dirFiles);
                    return;

                case SaveFormat.Mhtml: // ToDo: Check results of this assert

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.AllDirectories);
                    Assert.IsEmpty(dirFiles);
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

            doc.Save(ArtifactsDir + "Document.ExportListLabels.html", saveOptions);
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportUrlForLinkedImage(bool export)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportUrlForLinkedImage.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportOriginalUrlForLinkedImages = export };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

            string[] dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportUrlForLinkedImage.001.png", SearchOption.AllDirectories);

            if (dirFiles.Length == 0)
                DocumentHelper.FindTextInFile(ArtifactsDir + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
            else
                DocumentHelper.FindTextInFile(ArtifactsDir + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
        }

        [Test]
        public void ExportRoundtripInformation()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportRoundtripInformation = true };
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.RoundtripInformation.html", saveOptions);
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

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                ResourceFolder = "Resources",
                ResourceFolderAlias = "https://www.aspose.com/"
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportPageMargins.html", saveOptions);

            string[] imageFiles = Directory.GetFiles(ArtifactsDir + "Resources/", "*.png", SearchOption.AllDirectories);
            Assert.AreEqual(3, imageFiles.Length);

            string[] fontFiles = Directory.GetFiles(ArtifactsDir + "Resources/", "*.ttf", SearchOption.AllDirectories);
            Assert.AreEqual(1, fontFiles.Length);

            string[] cssFiles = Directory.GetFiles(ArtifactsDir + "Resources/", "*.css", SearchOption.AllDirectories);
            Assert.AreEqual(1, cssFiles.Length);

            DocumentHelper.FindTextInFile(ArtifactsDir + "HtmlSaveOptions.ExportPageMargins.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExportPageMargins.css\"");
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
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportPageMargins.html", saveOptions);
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

            doc.Save(ArtifactsDir + "HtmlSaveOptions.Html5Support.html", saveOptions);
        }

#if !__MOBILE__
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

                    doc.Save(ArtifactsDir + "DocumentExportFonts 1.html", saveOptions);
                    Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir, "DocumentExportFonts 1.times.ttf",
                        SearchOption.AllDirectories));
                    break;

                case true:

                    doc.Save(ArtifactsDir + "DocumentExportFonts 2.html", saveOptions);
                    Assert.IsEmpty(Directory.GetFiles(ArtifactsDir, "DocumentExportFonts 2.times.ttf",
                        SearchOption.AllDirectories));
                    break;
            }
        }
#endif

#if !__MOBILE__
        [Test]
        public void ResourceFolderPriority()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ResourceFolder.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.ResourceFolder = ArtifactsDir + "Resources";
            saveOptions.ResourceFolderAlias = "http://example.com/resources";

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ResourceFolder.html", saveOptions);

            string[] a = Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolder.001.jpeg",
                SearchOption.AllDirectories);
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolder.001.jpeg", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolder.002.png", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolder.calibri.ttf", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolder.css", SearchOption.AllDirectories));
        }
#endif

#if !__MOBILE__
        [Test]
        public void ResourceFolderLowPriority()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ResourceFolder.docx");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                FontsFolder = ArtifactsDir + "Fonts",
                ImagesFolder = ArtifactsDir + "Images",
                ResourceFolder = ArtifactsDir + "Resources",
                ResourceFolderAlias = "http://example.com/resources"
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ResourceFolder.html", saveOptions);

            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Images",
                "HtmlSaveOptions.ResourceFolder.001.jpeg", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Images", "HtmlSaveOptions.ResourceFolder.002.png",
                SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Fonts",
                "HtmlSaveOptions.ResourceFolder.calibri.ttf", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolder.css",
                SearchOption.AllDirectories));
        }
#endif

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

            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.MetafileFormat.html",
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

            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.MetafileFormat.html",
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

            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.MetafileFormat.html",
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

            doc.Save(ArtifactsDir + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
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

            doc.Save(ArtifactsDir + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
        }

        [Test]
        public void ContentIdScheme()
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ContentIdScheme.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                PrettyFormat = true,
                ExportCidUrlsForMhtmlResources = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ContentIdScheme.mhtml", saveOptions);
        }

        [Test]
        public void ResolveFontNames()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ResolveFontNames
            //ExSummary:Shows how to resolve all font names before writing them to HTML.
            Document document = new Document(MyDir + "HtmlSaveOptions.ResolveFontNames.docx");

            FontSettings fontSettings = new FontSettings
            {
                DefaultFontName = "Arial", 
                EnableFontSubstitution = false
            };

            document.FontSettings = fontSettings;

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // By default this option is set to 'False' and Aspose.Words writes font names as specified in the source document.
                ResolveFontNames = true 
            };

            document.Save(ArtifactsDir + "HtmlSaveOptions.ResolveFontNames.html", saveOptions);
            //ExEnd

            DocumentHelper.FindTextInFile(ArtifactsDir + "HtmlSaveOptions.ResolveFontNames.html", "<span style=\"font-family:Arial\">");
        }
    }
}