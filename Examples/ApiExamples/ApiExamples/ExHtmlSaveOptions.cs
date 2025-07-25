﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Fonts;
using Aspose.Words.Layout;
using Aspose.Words.Lists;
using Aspose.Words.Loading;
using NUnit.Framework;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
#if NET5_0_OR_GREATER || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        [TestCase(SaveFormat.Azw3)]
        [TestCase(SaveFormat.Mobi)]
        public void ExportPageMarginsEpub(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "TextBoxes.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                SaveFormat = saveFormat,
                ExportPageMargins = true
            };

            doc.Save(
                ArtifactsDir + "HtmlSaveOptions.ExportPageMarginsEpub" +
                FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);
        }

        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.Image)]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.MathML)]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.Text)]
        [TestCase(SaveFormat.Azw3, HtmlOfficeMathOutputMode.Text)]
        [TestCase(SaveFormat.Mobi, HtmlOfficeMathOutputMode.Text)]
        public void ExportOfficeMathEpub(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
        {
            Document doc = new Document(MyDir + "Office math.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { OfficeMathOutputMode = outputMode };

            doc.Save(
                ArtifactsDir + "HtmlSaveOptions.ExportOfficeMathEpub" +
                FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);
        }

        [TestCase(SaveFormat.Html, true, Description = "TextBox as svg (html)")]
        [TestCase(SaveFormat.Epub, true, Description = "TextBox as svg (epub)")]
        [TestCase(SaveFormat.Mhtml, false, Description = "TextBox as img (mhtml)")]
        [TestCase(SaveFormat.Azw3, false, Description = "TextBox as img (azw3)")]
        [TestCase(SaveFormat.Mobi, false, Description = "TextBox as img (mobi)")]
        public void ExportTextBoxAsSvgEpub(SaveFormat saveFormat, bool isTextBoxAsSvg)
        {
            string[] dirFiles;

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textbox = builder.InsertShape(ShapeType.TextBox, 300, 100);
            builder.MoveTo(textbox.FirstParagraph);
            builder.Write("Hello world!");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(saveFormat);
            saveOptions.ExportShapesAsSvg = isTextBoxAsSvg;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTextBoxAsSvgEpub" + FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);

            switch (saveFormat)
            {
                case SaveFormat.Html:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                        SearchOption.AllDirectories);
                    Assert.That(dirFiles.Length, Is.EqualTo(0));
                    return;

                case SaveFormat.Epub:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                        SearchOption.AllDirectories);
                    Assert.That(dirFiles.Length, Is.EqualTo(0));
                    return;

                case SaveFormat.Mhtml:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                        SearchOption.AllDirectories);
                    Assert.That(dirFiles.Length, Is.EqualTo(0));
                    return;

                case SaveFormat.Azw3:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                        SearchOption.AllDirectories);
                    Assert.That(dirFiles.Length, Is.EqualTo(0));
                    return;

                case SaveFormat.Mobi:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                        SearchOption.AllDirectories);
                    Assert.That(dirFiles.Length, Is.EqualTo(0));
                    return;
            }
        }

        [Test]
        public void CreateAZW3Toc()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.NavigationMapLevel
            //ExSummary:Shows how to generate table of contents for Azw3 documents.
            Document doc = new Document(MyDir + "Big document.docx");

            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Azw3);
            options.NavigationMapLevel = 2;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.CreateAZW3Toc.azw3", options);
            //ExEnd
        }

        [Test]
        public void CreateMobiToc()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.NavigationMapLevel
            //ExSummary:Shows how to generate table of contents for Mobi documents.
            Document doc = new Document(MyDir + "Big document.docx");

            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mobi);
            options.NavigationMapLevel = 5;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.CreateMobiToc.mobi", options);
            //ExEnd
        }

        [TestCase(ExportListLabels.Auto)]
        [TestCase(ExportListLabels.AsInlineText)]
        [TestCase(ExportListLabels.ByHtmlTags)]
        public void ControlListLabelsExport(ExportListLabels howExportListLabels)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.Lists.List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);
            builder.ListFormat.List = bulletedList;
            builder.ParagraphFormat.LeftIndent = 72;
            builder.Writeln("Bulleted list item 1.");
            builder.Writeln("Bulleted list item 2.");
            builder.ParagraphFormat.ClearFormatting();

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it does not cause formatting loss, 
                // otherwise HTML <p> tag is used. This is also the default value.
                // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation.
                // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
                ExportListLabels = howExportListLabels
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ControlListLabelsExport.html", saveOptions);
        }

        [TestCase(true)]
        [TestCase(false)]
        public void ExportUrlForLinkedImage(bool export)
        {
            Document doc = new Document(MyDir + "Linked image.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportOriginalUrlForLinkedImages = export };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

            string[] dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportUrlForLinkedImage.001.png",
                SearchOption.AllDirectories);

            DocumentHelper.FindTextInFile(ArtifactsDir + "HtmlSaveOptions.ExportUrlForLinkedImage.html",
                dirFiles.Length == 0
                    ? "<img src=\"http://www.aspose.com/images/aspose-logo.gif\""
                    : "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
        }

        [Test]
        public void ExportRoundtripInformation()
        {
            Document doc = new Document(MyDir + "TextBoxes.docx");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportRoundtripInformation = true };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.RoundtripInformation.html", saveOptions);
        }

        [Test]
        public void RoundtripInformationDefaulValue()
        {
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            Assert.That(saveOptions.ExportRoundtripInformation, Is.EqualTo(true));

            saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            Assert.That(saveOptions.ExportRoundtripInformation, Is.EqualTo(false));

            saveOptions = new HtmlSaveOptions(SaveFormat.Epub);
            Assert.That(saveOptions.ExportRoundtripInformation, Is.EqualTo(false));
        }

        [Test]
        public void ExternalResourceSavingConfig()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                ResourceFolder = "Resources",
                ResourceFolderAlias = "https://www.aspose.com/"
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExternalResourceSavingConfig.html", saveOptions);

            string[] imageFiles = Directory.GetFiles(ArtifactsDir + "Resources/",
                "HtmlSaveOptions.ExternalResourceSavingConfig*.png", SearchOption.AllDirectories);
            Assert.That(imageFiles.Length, Is.EqualTo(8));

            string[] fontFiles = Directory.GetFiles(ArtifactsDir + "Resources/",
                "HtmlSaveOptions.ExternalResourceSavingConfig*.ttf", SearchOption.AllDirectories);
            Assert.That(fontFiles.Length, Is.EqualTo(10));

            string[] cssFiles = Directory.GetFiles(ArtifactsDir + "Resources/",
                "HtmlSaveOptions.ExternalResourceSavingConfig*.css", SearchOption.AllDirectories);
            Assert.That(cssFiles.Length, Is.EqualTo(1));

            DocumentHelper.FindTextInFile(ArtifactsDir + "HtmlSaveOptions.ExternalResourceSavingConfig.html",
                "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExternalResourceSavingConfig.css\"");
        }

        [Test]
        public void ConvertFontsAsBase64()
        {
            Document doc = new Document(MyDir + "TextBoxes.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ResourceFolder = "Resources",
                ExportFontResources = true,
                ExportFontsAsBase64 = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ConvertFontsAsBase64.html", saveOptions);
        }

        [TestCase(HtmlVersion.Html5)]
        [TestCase(HtmlVersion.Xhtml)]
        public void Html5Support(HtmlVersion htmlVersion)
        {
            Document doc = new Document(MyDir + "Document.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { HtmlVersion = htmlVersion };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.Html5Support.html", saveOptions);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportFonts(bool exportAsBase64)
        {
            string fontsFolder = ArtifactsDir + "HtmlSaveOptions.ExportFonts.Resources";

            Document doc = new Document(MyDir + "Document.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,
                FontsFolder = fontsFolder,
                ExportFontsAsBase64 = exportAsBase64
            };

            switch (exportAsBase64)
            {
                case false:

                    doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportFonts.False.html", saveOptions);

                    Assert.That(Directory.GetFiles(fontsFolder, "HtmlSaveOptions.ExportFonts.False.times.ttf",
                        SearchOption.AllDirectories), Is.Not.Empty);

                    Directory.Delete(fontsFolder, true);
                    break;

                case true:

                    doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportFonts.True.html", saveOptions);
                    Assert.That(Directory.Exists(fontsFolder), Is.False);
                    break;
            }
        }

        [Test]
        public void ResourceFolderPriority()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                ResourceFolder = ArtifactsDir + "Resources",
                ResourceFolderAlias = "http://example.com/resources"
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ResourceFolderPriority.html", saveOptions);

            Assert.That(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.001.png", SearchOption.AllDirectories), Is.Not.Empty);
            Assert.That(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.002.png", SearchOption.AllDirectories), Is.Not.Empty);
            Assert.That(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.arial.ttf", SearchOption.AllDirectories), Is.Not.Empty);
            Assert.That(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.css", SearchOption.AllDirectories), Is.Not.Empty);
        }

        [Test]
        public void ResourceFolderLowPriority()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                FontsFolder = ArtifactsDir + "Fonts",
                ImagesFolder = ArtifactsDir + "Images",
                ResourceFolder = ArtifactsDir + "Resources",
                ResourceFolderAlias = "http://example.com/resources"
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ResourceFolderLowPriority.html", saveOptions);

            Assert.That(Directory.GetFiles(ArtifactsDir + "Images",
                "HtmlSaveOptions.ResourceFolderLowPriority.001.png", SearchOption.AllDirectories), Is.Not.Empty);
            Assert.That(Directory.GetFiles(ArtifactsDir + "Images", "HtmlSaveOptions.ResourceFolderLowPriority.002.png",
                SearchOption.AllDirectories), Is.Not.Empty);
            Assert.That(Directory.GetFiles(ArtifactsDir + "Fonts",
                "HtmlSaveOptions.ResourceFolderLowPriority.arial.ttf", SearchOption.AllDirectories), Is.Not.Empty);
            Assert.That(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderLowPriority.css",
                SearchOption.AllDirectories), Is.Not.Empty);
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

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.Png };
            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.SvgMetafileFormat.html", saveOptions);
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

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.Png };
            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.PngMetafileFormat.html", saveOptions);
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

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.EmfOrWmf };
            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.EmfOrWmfMetafileFormat.html", saveOptions);
        }

        [Test]
        public void CssClassNamesPrefix()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.CssClassNamePrefix
            //ExSummary:Shows how to save a document to HTML, and add a prefix to all of its CSS class names.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                CssClassNamePrefix = "myprefix-"
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.CssClassNamePrefix.html");

            Assert.That(outDocContents.Contains("<p class=\"myprefix-Header\">"), Is.True);
            Assert.That(outDocContents.Contains("<p class=\"myprefix-Footer\">"), Is.True);

            outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.CssClassNamePrefix.css");

            Assert.That(outDocContents.Contains(".myprefix-Footer { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt; -aw-style-name:footer }"), Is.True);
            Assert.That(outDocContents.Contains(".myprefix-Header { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt; -aw-style-name:header }"), Is.True);
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
            Document doc = new Document(MyDir + "Paragraphs.docx");

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
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                PrettyFormat = true,
                ExportCidUrlsForMhtmlResources = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ContentIdScheme.mhtml", saveOptions);
        }

        [TestCase(false)]
        [TestCase(true)]
        [Ignore("Bug")]
        public void ResolveFontNames(bool resolveFontNames)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ResolveFontNames
            //ExSummary:Shows how to resolve all font names before writing them to HTML.
            Document doc = new Document(MyDir + "Missing font.docx");

            // This document contains text that names a font that we do not have.
            Assert.That(doc.FontInfos["28 Days Later"], Is.Not.Null);

            // If we have no way of getting this font, and we want to be able to display all the text
            // in this document in an output HTML, we can substitute it with another font.
            FontSettings fontSettings = new FontSettings
            {
                SubstitutionSettings =
                {
                    DefaultFontSubstitution =
                    {
                        DefaultFontName = "Arial",
                        Enabled = true
                    }
                }
            };

            doc.FontSettings = fontSettings;

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // By default, this option is set to 'False' and Aspose.Words writes font names as specified in the source document
                ResolveFontNames = resolveFontNames
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ResolveFontNames.html", saveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ResolveFontNames.html");

            Assert.That(resolveFontNames
                ? Regex.Match(outDocContents, "<span style=\"font-family:Arial\">").Success
                : Regex.Match(outDocContents, "<span style=\"font-family:\'28 Days Later\'\">").Success, Is.True);
            //ExEnd
        }

        [Test]
        public void HeadingLevels()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.DocumentSplitHeadingLevel
            //ExSummary:Shows how to split an output HTML document by headings into several parts.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Every paragraph that we format using a "Heading" style can serve as a heading.
            // Each heading may also have a heading level, determined by the number of its heading style.
            // The headings below are of levels 1-3.
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 1"];
            builder.Writeln("Heading #1");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 2"];
            builder.Writeln("Heading #2");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 3"];
            builder.Writeln("Heading #3");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 1"];
            builder.Writeln("Heading #4");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 2"];
            builder.Writeln("Heading #5");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 3"];
            builder.Writeln("Heading #6");

            // Create a HtmlSaveOptions object and set the split criteria to "HeadingParagraph".
            // These criteria will split the document at paragraphs with "Heading" styles into several smaller documents,
            // and save each document in a separate HTML file in the local file system.
            // We will also set the maximum heading level, which splits the document to 2.
            // Saving the document will split it at headings of levels 1 and 2, but not at 3 to 9.
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 2
            };

            // Our document has four headings of levels 1 - 2. One of those headings will not be
            // a split point since it is at the beginning of the document.
            // The saving operation will split our document at three places, into four smaller documents.
            doc.Save(ArtifactsDir + "HtmlSaveOptions.HeadingLevels.html", options);

            doc = new Document(ArtifactsDir + "HtmlSaveOptions.HeadingLevels.html");

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Heading #1"));

            doc = new Document(ArtifactsDir + "HtmlSaveOptions.HeadingLevels-01.html");

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Heading #2\r" +
                            "Heading #3"));

            doc = new Document(ArtifactsDir + "HtmlSaveOptions.HeadingLevels-02.html");

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Heading #4"));

            doc = new Document(ArtifactsDir + "HtmlSaveOptions.HeadingLevels-03.html");

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Heading #5\r" +
                            "Heading #6"));
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void NegativeIndent(bool allowNegativeIndent)
        {
            //ExStart
            //ExFor:HtmlElementSizeOutputMode
            //ExFor:HtmlSaveOptions.AllowNegativeIndent
            //ExFor:HtmlSaveOptions.TableWidthOutputMode
            //ExSummary:Shows how to preserve negative indents in the output .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with a negative indent, which will push it to the left past the left page boundary.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndTable();
            table.LeftIndent = -36;
            table.PreferredWidth = PreferredWidth.FromPoints(144);

            builder.InsertBreak(BreakType.ParagraphBreak);

            // Insert a table with a positive indent, which will push the table to the right.
            table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndTable();
            table.LeftIndent = 36;
            table.PreferredWidth = PreferredWidth.FromPoints(144);

            // When we save a document to HTML, Aspose.Words will only preserve negative indents
            // such as the one we have applied to the first table if we set the "AllowNegativeIndent" flag
            // in a SaveOptions object that we will pass to "true".
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                AllowNegativeIndent = allowNegativeIndent,
                TableWidthOutputMode = HtmlElementSizeOutputMode.RelativeOnly
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.NegativeIndent.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.NegativeIndent.html");

            if (allowNegativeIndent)
            {
                Assert.That(outDocContents.Contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:-41.65pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"), Is.True);
                Assert.That(outDocContents.Contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"), Is.True);
            }
            else
            {
                Assert.That(outDocContents.Contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"), Is.True);
                Assert.That(outDocContents.Contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"), Is.True);
            }
            //ExEnd
        }

        [Test]
        public void FolderAlias()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportOriginalUrlForLinkedImages
            //ExFor:HtmlSaveOptions.FontsFolder
            //ExFor:HtmlSaveOptions.FontsFolderAlias
            //ExFor:HtmlSaveOptions.ImageResolution
            //ExFor:HtmlSaveOptions.ImagesFolderAlias
            //ExFor:HtmlSaveOptions.ResourceFolder
            //ExFor:HtmlSaveOptions.ResourceFolderAlias
            //ExSummary:Shows how to set folders and folder aliases for externally saved resources that Aspose.Words will create when saving a document to HTML.
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions options = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                ImageResolution = 72,
                FontResourcesSubsettingSizeThreshold = 0,
                FontsFolder = ArtifactsDir + "Fonts",
                ImagesFolder = ArtifactsDir + "Images",
                ResourceFolder = ArtifactsDir + "Resources",
                FontsFolderAlias = "http://example.com/fonts",
                ImagesFolderAlias = "http://example.com/images",
                ResourceFolderAlias = "http://example.com/resources",
                ExportOriginalUrlForLinkedImages = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.FolderAlias.html", options);
            //ExEnd
        }

        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontResources
        //ExFor:HtmlSaveOptions.FontSavingCallback
        //ExFor:IFontSavingCallback
        //ExFor:IFontSavingCallback.FontSaving
        //ExFor:FontSavingArgs
        //ExFor:FontSavingArgs.Bold
        //ExFor:FontSavingArgs.Document
        //ExFor:FontSavingArgs.FontFamilyName
        //ExFor:FontSavingArgs.FontFileName
        //ExFor:FontSavingArgs.FontStream
        //ExFor:FontSavingArgs.IsExportNeeded
        //ExFor:FontSavingArgs.IsSubsettingNeeded
        //ExFor:FontSavingArgs.Italic
        //ExFor:FontSavingArgs.KeepFontStreamOpen
        //ExFor:FontSavingArgs.OriginalFileName
        //ExFor:FontSavingArgs.OriginalFileSize
        //ExSummary:Shows how to define custom logic for exporting fonts when saving to HTML.
        [Test] //ExSkip
        public void SaveExportedFonts()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            // Configure a SaveOptions object to export fonts to separate files.
            // Set a callback that will handle font saving in a custom manner.
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportFontResources = true,
                FontSavingCallback = new HandleFontSaving()
            };

            // The callback will export .ttf files and save them alongside the output document.
            doc.Save(ArtifactsDir + "HtmlSaveOptions.SaveExportedFonts.html", options);

            foreach (string fontFilename in Array.FindAll(Directory.GetFiles(ArtifactsDir), s => s.EndsWith(".ttf")))
                Console.WriteLine(fontFilename);

            Assert.That(Array.FindAll(Directory.GetFiles(ArtifactsDir), s => s.EndsWith(".ttf")).Length, Is.EqualTo(10)); //ExSkip
        }

        /// <summary>
        /// Prints information about exported fonts and saves them in the same local system folder as their output .html.
        /// </summary>
        public class HandleFontSaving : IFontSavingCallback
        {
            void IFontSavingCallback.FontSaving(FontSavingArgs args)
            {
                Console.Write($"Font:\t{args.FontFamilyName}");
                if (args.Bold) Console.Write(", bold");
                if (args.Italic) Console.Write(", italic");
                Console.WriteLine($"\nSource:\t{args.OriginalFileName}, {args.OriginalFileSize} bytes\n");

                // We can also access the source document from here.
                Assert.That(args.Document.OriginalFileName.EndsWith("Rendering.docx"), Is.True);

                Assert.That(args.IsExportNeeded, Is.True);
                Assert.That(args.IsSubsettingNeeded, Is.True);

                // There are two ways of saving an exported font.
                // 1 -  Save it to a local file system location:
                args.FontFileName = args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last();

                // 2 -  Save it to a stream:
                args.FontStream =
                    new FileStream(ArtifactsDir + args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last(), FileMode.Create);
                Assert.That(args.KeepFontStreamOpen, Is.False);
            }
        }
        //ExEnd

        [TestCase(HtmlVersion.Html5)]
        [TestCase(HtmlVersion.Xhtml)]
        public void HtmlVersions(HtmlVersion htmlVersion)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
            //ExFor:HtmlSaveOptions.HtmlVersion
            //ExFor:HtmlVersion
            //ExSummary:Shows how to save a document to a specific version of HTML.
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                HtmlVersion = htmlVersion,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.HtmlVersions.html", options);

            // Our HTML documents will have minor differences to be compatible with different HTML versions.
            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.HtmlVersions.html");

            switch (htmlVersion)
            {
                case HtmlVersion.Html5:
                    Assert.That(outDocContents.Contains("<a id=\"_Toc76372689\"></a>"), Is.True);
                    Assert.That(outDocContents.Contains("<a id=\"_Toc76372689\"></a>"), Is.True);
                    Assert.That(outDocContents.Contains("<table style=\"padding:0pt; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"), Is.True);
                    break;
                case HtmlVersion.Xhtml:
                    Assert.That(outDocContents.Contains("<a name=\"_Toc76372689\"></a>"), Is.True);
                    Assert.That(outDocContents.Contains("<ul type=\"disc\" style=\"margin:0pt; padding-left:0pt\">"), Is.True);
                    Assert.That(outDocContents.Contains("<table cellspacing=\"0\" cellpadding=\"0\" style=\"-aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\""), Is.True);
                    break;
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportXhtmlTransitional(bool showDoctypeDeclaration)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportXhtmlTransitional
            //ExFor:HtmlSaveOptions.HtmlVersion
            //ExFor:HtmlVersion
            //ExSummary:Shows how to display a DOCTYPE heading when converting documents to the Xhtml 1.0 transitional standard.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                HtmlVersion = HtmlVersion.Xhtml,
                ExportXhtmlTransitional = showDoctypeDeclaration,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportXhtmlTransitional.html", options);

            // Our document will only contain a DOCTYPE declaration heading if we have set the "ExportXhtmlTransitional" flag to "true".
            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportXhtmlTransitional.html");
            string newLine = Environment.NewLine;

            if (showDoctypeDeclaration)
                Assert.That(outDocContents.Contains(
                    $"<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>{newLine}" +
                    $"<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">{newLine}" +
                    "<html xmlns=\"http://www.w3.org/1999/xhtml\">"), Is.True);
            else
                Assert.That(outDocContents.Contains("<html>"), Is.True);
            //ExEnd
        }

        [Test]
        public void EpubHeadings()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.NavigationMapLevel
            //ExSummary:Shows how to filter headings that appear in the navigation panel of a saved Epub document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Every paragraph that we format using a "Heading" style can serve as a heading.
            // Each heading may also have a heading level, determined by the number of its heading style.
            // The headings below are of levels 1-3.
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 1"];
            builder.Writeln("Heading #1");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 2"];
            builder.Writeln("Heading #2");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 3"];
            builder.Writeln("Heading #3");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 1"];
            builder.Writeln("Heading #4");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 2"];
            builder.Writeln("Heading #5");
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 3"];
            builder.Writeln("Heading #6");

            // Epub readers typically create a table of contents for their documents.
            // Each paragraph with a "Heading" style in the document will create an entry in this table of contents.
            // We can use the "NavigationMapLevel" property to set a maximum heading level. 
            // The Epub reader will not add headings with a level above the one we specify to the contents table.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Epub);
            options.NavigationMapLevel = 2;

            // Our document has six headings, two of which are above level 2.
            // The table of contents for this document will have four entries.
            doc.Save(ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", options);
            //ExEnd

            TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #1</text></navLabel>", 
                ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");
            TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #2</text></navLabel>", 
                ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");
            TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #4</text></navLabel>", 
                ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");
            TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #5</text></navLabel>", 
                ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");

            Assert.Throws<AssertionException>(() =>
            {
                TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #3</text></navLabel>", 
                    ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");
            });

            Assert.Throws<AssertionException>(() =>
            {
                TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #6</text></navLabel>", 
                    ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");
            });
        }

        [Test]
        public void Doc2EpubSaveOptions()
        {
            //ExStart
            //ExFor:DocumentSplitCriteria
            //ExFor:HtmlSaveOptions
            //ExFor:HtmlSaveOptions.#ctor
            //ExFor:HtmlSaveOptions.Encoding
            //ExFor:HtmlSaveOptions.DocumentSplitCriteria
            //ExFor:HtmlSaveOptions.ExportDocumentProperties
            //ExFor:HtmlSaveOptions.SaveFormat
            //ExFor:SaveOptions
            //ExFor:SaveOptions.SaveFormat
            //ExSummary:Shows how to use a specific encoding when saving a document to .epub.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Use a SaveOptions object to specify the encoding for a document that we will save.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Epub;
            saveOptions.Encoding = Encoding.UTF8;

            // By default, an output .epub document will have all its contents in one HTML part.
            // A split criterion allows us to segment the document into several HTML parts.
            // We will set the criteria to split the document into heading paragraphs.
            // This is useful for readers who cannot read HTML files more significant than a specific size.
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

            // Specify that we want to export document properties.
            saveOptions.ExportDocumentProperties = true;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.Doc2EpubSaveOptions.epub", saveOptions);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ContentIdUrls(bool exportCidUrlsForMhtmlResources)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportCidUrlsForMhtmlResources
            //ExSummary:Shows how to enable content IDs for output MHTML documents.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Setting this flag will replace "Content-Location" tags
            // with "Content-ID" tags for each resource from the input document.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportCidUrlsForMhtmlResources = exportCidUrlsForMhtmlResources,
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ContentIdUrls.mht", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ContentIdUrls.mht");

            if (exportCidUrlsForMhtmlResources)
            {
                Assert.That(outDocContents.Contains("Content-ID: <document.html>"), Is.True);
                Assert.That(outDocContents.Contains("<link href=3D\"cid:styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"), Is.True);
                Assert.That(outDocContents.Contains("@font-face { font-family:'Arial Black'; font-weight:bold; src:url('cid:arib=\r\nlk.ttf') }"), Is.True);
                Assert.That(outDocContents.Contains("<img src=3D\"cid:image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />"), Is.True);
            }
            else
            {
                Assert.That(outDocContents.Contains("Content-Location: document.html"), Is.True);
                Assert.That(outDocContents.Contains("<link href=3D\"styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"), Is.True);
                Assert.That(outDocContents.Contains("@font-face { font-family:'Arial Black'; font-weight:bold; src:url('ariblk.t=\r\ntf') }"), Is.True);
                Assert.That(outDocContents.Contains("<img src=3D\"image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />"), Is.True);
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void DropDownFormField(bool exportDropDownFormFieldAsText)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportDropDownFormFieldAsText
            //ExSummary:Shows how to get drop-down combo box form fields to blend in with paragraph text when saving to html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a combo box with the value "Two" selected.
            builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 1);

            // The "ExportDropDownFormFieldAsText" flag of this SaveOptions object allows us to
            // control how saving the document to HTML treats drop-down combo boxes.
            // Setting it to "true" will convert each combo box into simple text
            // that displays the combo box's currently selected value, effectively freezing it.
            // Setting it to "false" will preserve the functionality of the combo box using <select> and <option> tags.
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportDropDownFormFieldAsText = exportDropDownFormFieldAsText;    

            doc.Save(ArtifactsDir + "HtmlSaveOptions.DropDownFormField.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.DropDownFormField.html");

            if (exportDropDownFormFieldAsText)
                Assert.That(outDocContents.Contains(
                    "<span>Two</span>"), Is.True);
            else
                Assert.That(outDocContents.Contains(
                    "<select name=\"MyComboBox\">" +
                        "<option>One</option>" +
                        "<option selected=\"selected\">Two</option>" +
                        "<option>Three</option>" +
                    "</select>"), Is.True);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportImagesAsBase64(bool exportImagesAsBase64)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportFontsAsBase64
            //ExFor:HtmlSaveOptions.ExportImagesAsBase64
            //ExSummary:Shows how to save a .html document with images embedded inside it.
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = exportImagesAsBase64,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportImagesAsBase64.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportImagesAsBase64.html");

            Assert.That(exportImagesAsBase64
                ? outDocContents.Contains("<img src=\"data:image/png;base64")
                : outDocContents.Contains("<img src=\"HtmlSaveOptions.ExportImagesAsBase64.001.png\""), Is.True);
            //ExEnd
        }


        [Test]
        public void ExportFontsAsBase64()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportFontsAsBase64
            //ExFor:HtmlSaveOptions.ExportImagesAsBase64
            //ExSummary:Shows how to embed fonts inside a saved HTML document.
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportFontsAsBase64 = true,
                CssStyleSheetType = CssStyleSheetType.Embedded,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportFontsAsBase64.html", options);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportLanguageInformation(bool exportLanguageInformation)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportLanguageInformation
            //ExSummary:Shows how to preserve language information when saving to .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use the builder to write text while formatting it in different locales.
            builder.Font.LocaleId = new CultureInfo("en-US").LCID;
            builder.Writeln("Hello world!");

            builder.Font.LocaleId = new CultureInfo("en-GB").LCID;
            builder.Writeln("Hello again!");

            builder.Font.LocaleId = new CultureInfo("ru-RU").LCID;
            builder.Write("Привет, мир!");

            // When saving the document to HTML, we can pass a SaveOptions object
            // to either preserve or discard each formatted text's locale.
            // If we set the "ExportLanguageInformation" flag to "true",
            // the output HTML document will contain the locales in "lang" attributes of <span> tags.
            // If we set the "ExportLanguageInformation" flag to "false',
            // the text in the output HTML document will not contain any locale information.
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportLanguageInformation = exportLanguageInformation,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportLanguageInformation.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportLanguageInformation.html");

            if (exportLanguageInformation)
            {
                Assert.That(outDocContents.Contains("<span>Hello world!</span>"), Is.True);
                Assert.That(outDocContents.Contains("<span lang=\"en-GB\">Hello again!</span>"), Is.True);
                Assert.That(outDocContents.Contains("<span lang=\"ru-RU\">Привет, мир!</span>"), Is.True);
            }
            else
            {
                Assert.That(outDocContents.Contains("<span>Hello world!</span>"), Is.True);
                Assert.That(outDocContents.Contains("<span>Hello again!</span>"), Is.True);
                Assert.That(outDocContents.Contains("<span>Привет, мир!</span>"), Is.True);
            }
            //ExEnd
        }

        [TestCase(ExportListLabels.AsInlineText)]
        [TestCase(ExportListLabels.Auto)]
        [TestCase(ExportListLabels.ByHtmlTags)]
        public void List(ExportListLabels exportListLabels)
        {
            //ExStart
            //ExFor:ExportListLabels
            //ExFor:HtmlSaveOptions.ExportListLabels
            //ExSummary:Shows how to configure list exporting to HTML.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.Lists.List docList = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = docList;
            
            builder.Writeln("Default numbered list item 1.");
            builder.Writeln("Default numbered list item 2.");
            builder.ListFormat.ListIndent();
            builder.Writeln("Default numbered list item 3.");
            builder.ListFormat.RemoveNumbers();

            docList = doc.Lists.Add(ListTemplate.OutlineHeadingsLegal);
            builder.ListFormat.List = docList;

            builder.Writeln("Outline legal heading list item 1.");
            builder.Writeln("Outline legal heading list item 2.");
            builder.ListFormat.ListIndent();
            builder.Writeln("Outline legal heading list item 3.");
            builder.ListFormat.ListIndent();
            builder.Writeln("Outline legal heading list item 4.");
            builder.ListFormat.ListIndent();
            builder.Writeln("Outline legal heading list item 5.");
            builder.ListFormat.RemoveNumbers();

            // When saving the document to HTML, we can pass a SaveOptions object
            // to decide which HTML elements the document will use to represent lists.
            // Setting the "ExportListLabels" property to "ExportListLabels.AsInlineText"
            // will create lists by formatting spans.
            // Setting the "ExportListLabels" property to "ExportListLabels.Auto" will use the <p> tag
            // to build lists in cases when using the <ol> and <li> tags may cause loss of formatting.
            // Setting the "ExportListLabels" property to "ExportListLabels.ByHtmlTags"
            // will use <ol> and <li> tags to build all lists.
            HtmlSaveOptions options = new HtmlSaveOptions { ExportListLabels = exportListLabels };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.List.html", options);
            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.List.html");

            switch (exportListLabels)
            {
                case ExportListLabels.AsInlineText:
                    Assert.That(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-left:72pt; margin-bottom:0pt; text-indent:-18pt; -aw-import:list-item; -aw-list-level-number:1; -aw-list-number-format:'%1.'; -aw-list-number-styles:'lowerLetter'; -aw-list-number-values:'1'; -aw-list-padding-sml:9.67pt\">" +
                            "<span style=\"-aw-import:ignore\">" +
                                "<span>a.</span>" +
                                "<span style=\"width:9.67pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "</span>" +
                            "<span>Default numbered list item 3.</span>" +
                        "</p>"), Is.True);

                    Assert.That(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; -aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; -aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" +
                            "<span style=\"-aw-import:ignore\">" +
                                "<span>2.1.1.1</span>" +
                                "<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "</span>" +
                            "<span>Outline legal heading list item 5.</span>" +
                        "</p>"), Is.True);
                    break;
                case ExportListLabels.Auto:
                    Assert.That(outDocContents.Contains(
                        "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                            "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
                                "<span>Default numbered list item 3.</span>" +
                            "</li>" +
                        "</ol>"), Is.True);
                    break;
                case ExportListLabels.ByHtmlTags:
                    Assert.That(outDocContents.Contains(
                        "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                            "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
                                "<span>Default numbered list item 3.</span>" +
                            "</li>" +
                        "</ol>"), Is.True);
                    break;
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportPageMargins(bool exportPageMargins)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportPageMargins
            //ExSummary:Shows how to show out-of-bounds objects in output HTML documents.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a builder to insert a shape with no wrapping.
            Shape shape = builder.InsertShape(ShapeType.Cube, 200, 200);

            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            shape.WrapType = WrapType.None;

            // Negative shape position values may place the shape outside of page boundaries.
            // If we export this to HTML, the shape will appear truncated.
            shape.Left = -150;

            // When saving the document to HTML, we can pass a SaveOptions object
            // to decide whether to adjust the page to display out-of-bounds objects fully.
            // If we set the "ExportPageMargins" flag to "true", the shape will be fully visible in the output HTML.
            // If we set the "ExportPageMargins" flag to "false",
            // our document will display the shape truncated as we would see it in Microsoft Word.
            HtmlSaveOptions options = new HtmlSaveOptions { ExportPageMargins = exportPageMargins };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportPageMargins.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportPageMargins.html");

            if (exportPageMargins)
            {
                Assert.That(outDocContents.Contains("<style type=\"text/css\">div.Section_1 { margin:70.85pt }</style>"), Is.True);
                Assert.That(outDocContents.Contains("<div class=\"Section_1\"><p style=\"margin-top:0pt; margin-left:150pt; margin-bottom:0pt\">"), Is.True);
            }
            else
            {
                Assert.That(outDocContents.Contains("style type=\"text/css\">"), Is.False);
                Assert.That(outDocContents.Contains("<div><p style=\"margin-top:0pt; margin-left:220.85pt; margin-bottom:0pt\">"), Is.True);
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportPageSetup(bool exportPageSetup)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportPageSetup
            //ExSummary:Shows how decide whether to preserve section structure/page setup information when saving to HTML.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");

            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.TopMargin = 36.0;
            pageSetup.BottomMargin = 36.0;
            pageSetup.PaperSize = PaperSize.A5;

            // When saving the document to HTML, we can pass a SaveOptions object
            // to decide whether to preserve or discard page setup settings.
            // If we set the "ExportPageSetup" flag to "true", the output HTML document will contain our page setup configuration.
            // If we set the "ExportPageSetup" flag to "false", the save operation will discard our page setup settings
            // for the first section, and both sections will look identical.
            HtmlSaveOptions options = new HtmlSaveOptions { ExportPageSetup = exportPageSetup };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportPageSetup.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportPageSetup.html");

            if (exportPageSetup)
            {
                Assert.That(outDocContents.Contains(
                    "<style type=\"text/css\">" +
                        "@page Section_1 { size:419.55pt 595.3pt; margin:36pt 70.85pt; -aw-footer-distance:35.4pt; -aw-header-distance:35.4pt }" +
                        "@page Section_2 { size:612pt 792pt; margin:70.85pt; -aw-footer-distance:35.4pt; -aw-header-distance:35.4pt }" +
                        "div.Section_1 { page:Section_1 }div.Section_2 { page:Section_2 }" +
                    "</style>"), Is.True);

                Assert.That(outDocContents.Contains(
                    "<div class=\"Section_1\">" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Section 1</span>" +
                        "</p>" +
                    "</div>"), Is.True);
            }
            else
            {
                Assert.That(outDocContents.Contains("style type=\"text/css\">"), Is.False);

                Assert.That(outDocContents.Contains(
                    "<div>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Section 1</span>" +
                        "</p>" +
                    "</div>"), Is.True);
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void RelativeFontSize(bool exportRelativeFontSize)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportRelativeFontSize
            //ExSummary:Shows how to use relative font sizes when saving to .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Default font size, ");
            builder.Font.Size = 24;
            builder.Writeln("2x default font size,");
            builder.Font.Size = 96;
            builder.Write("8x default font size");

            // When we save the document to HTML, we can pass a SaveOptions object
            // to determine whether to use relative or absolute font sizes.
            // Set the "ExportRelativeFontSize" flag to "true" to declare font sizes
            // using the "em" measurement unit, which is a factor that multiplies the current font size. 
            // Set the "ExportRelativeFontSize" flag to "false" to declare font sizes
            // using the "pt" measurement unit, which is the font's absolute size in points.
            HtmlSaveOptions options = new HtmlSaveOptions { ExportRelativeFontSize = exportRelativeFontSize };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.RelativeFontSize.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.RelativeFontSize.html");

            if (exportRelativeFontSize)
            {
                Assert.That(outDocContents.Contains(
                    "<body style=\"font-family:'Times New Roman'\">" +
                        "<div>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                                "<span>Default font size, </span>" +
                            "</p>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:2em\">" +
                                "<span>2x default font size,</span>" +
                            "</p>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:8em\">" +
                                "<span>8x default font size</span>" +
                            "</p>" +
                        "</div>" +
                    "</body>"), Is.True);
            }
            else
            {
                Assert.That(outDocContents.Contains(
                    "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                        "<div>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                                "<span>Default font size, </span>" +
                            "</p>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:24pt\">" +
                                "<span>2x default font size,</span>" +
                            "</p>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:96pt\">" +
                                "<span>8x default font size</span>" +
                            "</p>" +
                        "</div>" +
                    "</body>"), Is.True);
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportShape(bool exportShapesAsSvg)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportShapesAsSvg
            //ExSummary:Shows how to export shape as scalable vector graphics.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100.0, 60.0);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("My text box");

            // When we save the document to HTML, we can pass a SaveOptions object
            // to determine how the saving operation will export text box shapes.
            // If we set the "ExportTextBoxAsSvg" flag to "true",
            // the save operation will convert shapes with text into SVG objects.
            // If we set the "ExportTextBoxAsSvg" flag to "false",
            // the save operation will convert shapes with text into images.
            HtmlSaveOptions options = new HtmlSaveOptions { ExportShapesAsSvg = exportShapesAsSvg };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTextBox.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportTextBox.html");

            if (exportShapesAsSvg)
            {
                Assert.That(outDocContents.Contains(
                    "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"133\" height=\"80\">"), Is.True);
            }
            else
            {
                Assert.That(outDocContents.Contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                        "<img src=\"HtmlSaveOptions.ExportTextBox.001.png\" width=\"136\" height=\"83\" alt=\"\" " +
                        "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                    "</p>"), Is.True);
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void RoundTripInformation(bool exportRoundtripInformation)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportRoundtripInformation
            //ExSummary:Shows how to preserve hidden elements when converting to .html.
            Document doc = new Document(MyDir + "Rendering.docx");

            // When converting a document to .html, some elements such as hidden bookmarks, original shape positions,
            // or footnotes will be either removed or converted to plain text and effectively be lost.
            // Saving with a HtmlSaveOptions object with ExportRoundtripInformation set to true will preserve these elements.

            // When we save the document to HTML, we can pass a SaveOptions object to determine
            // how the saving operation will export document elements that HTML does not support or use,
            // such as hidden bookmarks and original shape positions.
            // If we set the "ExportRoundtripInformation" flag to "true", the save operation will preserve these elements.
            // If we set the "ExportRoundTripInformation" flag to "false", the save operation will discard these elements.
            // We will want to preserve such elements if we intend to load the saved HTML using Aspose.Words,
            // as they could be of use once again.
            HtmlSaveOptions options = new HtmlSaveOptions { ExportRoundtripInformation = exportRoundtripInformation };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.RoundTripInformation.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.RoundTripInformation.html");
            doc = new Document(ArtifactsDir + "HtmlSaveOptions.RoundTripInformation.html");

            if (exportRoundtripInformation)
            {
                Assert.That(outDocContents.Contains("<div style=\"-aw-headerfooter-type:header-primary; clear:both\">"), Is.True);
                Assert.That(outDocContents.Contains("<span style=\"-aw-import:ignore\">&#xa0;</span>"), Is.True);

                Assert.That(outDocContents.Contains(
                    "td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                    "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top; " +
                    "-aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single; -aw-border-top:0.5pt single\">"), Is.True);

                Assert.That(outDocContents.Contains(
                    "<li style=\"margin-left:30.2pt; padding-left:5.8pt; -aw-font-family:'Courier New'; -aw-font-weight:normal; -aw-number-format:'o'\">"), Is.True);

                Assert.That(outDocContents.Contains(
                    "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" " +
                    "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />"), Is.True);


                Assert.That(outDocContents.Contains(
                    "<span>Page number </span>" +
                    "<span style=\"-aw-field-start:true\"></span>" +
                    "<span style=\"-aw-field-code:' PAGE   \\\\* MERGEFORMAT '\"></span>" +
                    "<span style=\"-aw-field-separator:true\"></span>" +
                    "<span>1</span>" +
                    "<span style=\"-aw-field-end:true\"></span>"), Is.True);

                Assert.That(doc.Range.Fields.Count(f => f.Type == FieldType.FieldPage), Is.EqualTo(1));
            }
            else
            {
                Assert.That(outDocContents.Contains("<div style=\"clear:both\">"), Is.True);
                Assert.That(outDocContents.Contains("<span>&#xa0;</span>"), Is.True);

                Assert.That(outDocContents.Contains(
                    "<td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                    "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top\">"), Is.True);
                
                Assert.That(outDocContents.Contains(
                    "<li style=\"margin-left:30.2pt; padding-left:5.8pt\">"), Is.True);

                Assert.That(outDocContents.Contains(
                    "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" />"), Is.True);

                Assert.That(outDocContents.Contains(
                    "<span>Page number 1</span>"), Is.True);

                Assert.That(doc.Range.Fields.Count(f => f.Type == FieldType.FieldPage), Is.EqualTo(0));
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportTocPageNumbers(bool exportTocPageNumbers)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportTocPageNumbers
            //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents, and then populate the document with paragraphs formatted using a "Heading"
            // style that the table of contents will pick up as entries. Each entry will display the heading paragraph on the left,
            // and the page number that contains the heading on the right.
            FieldToc fieldToc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 1"];
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Entry 1");
            builder.Writeln("Entry 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Entry 3");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Entry 4");
            fieldToc.UpdatePageNumbers();
            doc.UpdateFields();

            // HTML documents do not have pages. If we save this document to HTML,
            // the page numbers that our TOC displays will have no meaning.
            // When we save the document to HTML, we can pass a SaveOptions object to omit these page numbers from the TOC.
            // If we set the "ExportTocPageNumbers" flag to "true",
            // each TOC entry will display the heading, separator, and page number, preserving its appearance in Microsoft Word.
            // If we set the "ExportTocPageNumbers" flag to "false",
            // the save operation will omit both the separator and page number and leave the heading for each entry intact.
            HtmlSaveOptions options = new HtmlSaveOptions { ExportTocPageNumbers = exportTocPageNumbers };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTocPageNumbers.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportTocPageNumbers.html");

            if (exportTocPageNumbers)
            {
                Assert.That(outDocContents.Contains(
                    "<span>Entry 1</span>" +
                    "<span style=\"width:428.14pt; font-family:'Lucida Console'; font-size:10pt; display:inline-block; -aw-font-family:'Times New Roman'; " +
                    "-aw-tabstop-align:right; -aw-tabstop-leader:dots; -aw-tabstop-pos:469.8pt\">.......................................................................</span>" +
                    "<span>2</span>" +
                    "</p>"), Is.True);
            }
            else
            {
                Assert.That(outDocContents.Contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                    "<span>Entry 2</span>" +
                    "</p>"), Is.True);
            }
            //ExEnd
        }

        [TestCase(0)]
        [TestCase(1000000)]
        [TestCase(int.MaxValue)]
        public void FontSubsetting(int fontResourcesSubsettingSizeThreshold)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.FontResourcesSubsettingSizeThreshold
            //ExSummary:Shows how to work with font subsetting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Times New Roman";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Courier New";
            builder.Writeln("Hello world!");

            // When we save the document to HTML, we can pass a SaveOptions object configure font subsetting.
            // Suppose we set the "ExportFontResources" flag to "true" and also name a folder in the "FontsFolder" property.
            // In that case, the saving operation will create that folder and place a .ttf file inside
            // that folder for each font that our document uses.
            // Each .ttf file will contain that font's entire glyph set,
            // which may potentially result in a very large file that accompanies the document.
            // When we apply subsetting to a font, its exported raw data will only contain the glyphs that the document is
            // using instead of the entire glyph set. If the text in our document only uses a small fraction of a font's
            // glyph set, then subsetting will significantly reduce our output documents' size.
            // We can use the "FontResourcesSubsettingSizeThreshold" property to define a .ttf file size, in bytes.
            // If an exported font creates a size bigger file than that, then the save operation will apply subsetting to that font. 
            // Setting a threshold of 0 applies subsetting to all fonts,
            // and setting it to "int.MaxValue" effectively disables subsetting.
            string fontsFolder = ArtifactsDir + "HtmlSaveOptions.FontSubsetting.Fonts";

            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportFontResources = true,
                FontsFolder = fontsFolder,
                FontResourcesSubsettingSizeThreshold = fontResourcesSubsettingSizeThreshold
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.FontSubsetting.html", options);

            string[] fontFileNames = Directory.GetFiles(fontsFolder).Where(s => s.EndsWith(".ttf")).ToArray();

            Assert.That(fontFileNames.Length, Is.EqualTo(3));

            foreach (string filename in fontFileNames)
            {
                // By default, the .ttf files for each of our three fonts will be over 700MB.
                // Subsetting will reduce them all to under 30MB.
                FileInfo fontFileInfo = new FileInfo(filename);

                Assert.That(fontFileInfo.Length > 700000 || fontFileInfo.Length < 30000, Is.True);
                Assert.That(System.Math.Max(fontResourcesSubsettingSizeThreshold, 30000) > new FileInfo(filename).Length, Is.True);
            }
            //ExEnd
        }

        [TestCase(HtmlMetafileFormat.Png)]
        [TestCase(HtmlMetafileFormat.Svg)]
        [TestCase(HtmlMetafileFormat.EmfOrWmf)]
        public void MetafileFormat(HtmlMetafileFormat htmlMetafileFormat)
        {
            //ExStart
            //ExFor:HtmlMetafileFormat
            //ExFor:HtmlSaveOptions.MetafileFormat
            //ExFor:HtmlLoadOptions.ConvertSvgToEmf
            //ExSummary:Shows how to convert SVG objects to a different format when saving HTML documents.
            string html = 
                @"<html>
                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>
                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>
                    </svg>
                </html>";

            // Use 'ConvertSvgToEmf' to turn back the legacy behavior
            // where all SVG images loaded from an HTML document were converted to EMF.
            // Now SVG images are loaded without conversion
            // if the MS Word version specified in load options supports SVG images natively.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions { ConvertSvgToEmf = true };

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), loadOptions);

            // This document contains a <svg> element in the form of text.
            // When we save the document to HTML, we can pass a SaveOptions object
            // to determine how the saving operation handles this object.
            // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Png" to convert it to a PNG image.
            // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Svg" preserve it as a SVG object.
            // Setting the "MetafileFormat" property to "HtmlMetafileFormat.EmfOrWmf" to convert it to a metafile.
            HtmlSaveOptions options = new HtmlSaveOptions { MetafileFormat = htmlMetafileFormat };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.MetafileFormat.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.MetafileFormat.html");

            switch (htmlMetafileFormat)
            {
                case HtmlMetafileFormat.Png:
                    Assert.That(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.MetafileFormat.001.png\" width=\"500\" height=\"40\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>"), Is.True);
                    break;
                case HtmlMetafileFormat.Svg:
                    Assert.That(outDocContents.Contains(
                        "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                        "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"499\" height=\"40\">"), Is.True);
                    break;
                case HtmlMetafileFormat.EmfOrWmf:
                    Assert.That(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.MetafileFormat.001.emf\" width=\"500\" height=\"40\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>"), Is.True);
                    break;
            }
            //ExEnd
        }

        [TestCase(HtmlOfficeMathOutputMode.Image)]
        [TestCase(HtmlOfficeMathOutputMode.MathML)]
        [TestCase(HtmlOfficeMathOutputMode.Text)]
        public void OfficeMathOutputMode(HtmlOfficeMathOutputMode htmlOfficeMathOutputMode)
        {
            //ExStart
            //ExFor:HtmlOfficeMathOutputMode
            //ExFor:HtmlSaveOptions.OfficeMathOutputMode
            //ExSummary:Shows how to specify how to export Microsoft OfficeMath objects to HTML.
            Document doc = new Document(MyDir + "Office math.docx");

            // When we save the document to HTML, we can pass a SaveOptions object
            // to determine how the saving operation handles OfficeMath objects.
            // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Image"
            // will render each OfficeMath object into an image.
            // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.MathML"
            // will convert each OfficeMath object into MathML.
            // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Text"
            // will represent each OfficeMath formula using plain HTML text.
            HtmlSaveOptions options = new HtmlSaveOptions { OfficeMathOutputMode = htmlOfficeMathOutputMode };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.OfficeMathOutputMode.html", options);
            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.OfficeMathOutputMode.html");

            switch (htmlOfficeMathOutputMode)
            {
                case HtmlOfficeMathOutputMode.Image:
                    Assert.That(Regex.Match(outDocContents,
                        "<p style=\"margin-top:0pt; margin-bottom:10pt\">" +
                            "<img src=\"HtmlSaveOptions.OfficeMathOutputMode.001.png\" width=\"163\" height=\"19\" alt=\"\" style=\"vertical-align:middle; " +
                            "-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>").Success, Is.True);
                    break;
                case HtmlOfficeMathOutputMode.MathML:
                    Assert.That(Regex.Match(outDocContents,
                        "<p style=\"margin-top:0pt; margin-bottom:10pt; text-align:center\">" +
                            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                                "<mi>i</mi>" +
                                "<mo>[+]</mo>" +
                                "<mi>b</mi>" +
                                "<mo>-</mo>" +
                                "<mi>c</mi>" +
                                "<mo>≥</mo>" +
                                ".*" +
                            "</math>" +
                        "</p>").Success, Is.True);
                    break;
                case HtmlOfficeMathOutputMode.Text:
                    Assert.That(Regex.Match(outDocContents,
                        @"<p style=\""margin-top:0pt; margin-bottom:10pt; text-align:center\"">" +
                            @"<span style=\""font-family:'Cambria Math'\"">i[+]b-c≥iM[+]bM-cM </span>" +
                        "</p>").Success, Is.True);
                    break;
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ScaleImageToShapeSize(bool scaleImageToShapeSize)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ScaleImageToShapeSize
            //ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a shape which contains an image, and then make that shape considerably smaller than the image.
            Shape imageShape = builder.InsertImage(ImageDir + "Transparent background logo.png");
            imageShape.Width = 50;
            imageShape.Height = 50;

            // Saving a document that contains shapes with images to HTML will create an image file in the local file system
            // for each such shape. The output HTML document will use <image> tags to link to and display these images.
            // When we save the document to HTML, we can pass a SaveOptions object to determine
            // whether to scale all images that are inside shapes to the sizes of their shapes.
            // Setting the "ScaleImageToShapeSize" flag to "true" will shrink every image
            // to the size of the shape that contains it, so that no saved images will be larger than the document requires them to be.
            // Setting the "ScaleImageToShapeSize" flag to "false" will preserve these images' original sizes,
            // which will take up more space in exchange for preserving image quality.
            HtmlSaveOptions options = new HtmlSaveOptions { ScaleImageToShapeSize = scaleImageToShapeSize };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ScaleImageToShapeSize.html", options);
            //ExEnd

            var testedImageLength = new FileInfo(ArtifactsDir + "HtmlSaveOptions.ScaleImageToShapeSize.001.png").Length;

            if (scaleImageToShapeSize)
#if NETFRAMEWORK || JAVA || CPLUSPLUS
                Assert.That(testedImageLength < 3000, Is.True);
#elif NET6_0_OR_GREATER
                Assert.That(testedImageLength < 6200, Is.True);
#endif
            else
                Assert.That(testedImageLength < 16000, Is.True);
            
        }

        [Test]
        public void ImageFolder()
        {
            //ExStart
            //ExFor:HtmlSaveOptions
            //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
            //ExFor:HtmlSaveOptions.ImagesFolder
            //ExSummary:Shows how to specify the folder for storing linked images after saving to .html.
            Document doc = new Document(MyDir + "Rendering.docx");

            string imagesDir = Path.Combine(ArtifactsDir, "SaveHtmlWithOptions");

            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text instead of HTML input elements.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportTextInputFormFieldAsText = true, 
                ImagesFolder = imagesDir
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.SaveHtmlWithOptions.html", options);
            //ExEnd

            Assert.That(File.Exists(ArtifactsDir + "HtmlSaveOptions.SaveHtmlWithOptions.html"), Is.True);
            Assert.That(Directory.GetFiles(imagesDir).Length, Is.EqualTo(9));

            Directory.Delete(imagesDir, true);
        }

        //ExStart
        //ExFor:ImageSavingArgs.CurrentShape
        //ExFor:ImageSavingArgs.Document
        //ExFor:ImageSavingArgs.ImageStream
        //ExFor:ImageSavingArgs.IsImageAvailable
        //ExFor:ImageSavingArgs.KeepImageStreamOpen
        //ExSummary:Shows how to involve an image saving callback in an HTML conversion process.
        [Test] //ExSkip
        public void ImageSavingCallback()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            // When we save the document to HTML, we can pass a SaveOptions object to designate a callback
            // to customize the image saving process.
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ImageSavingCallback = new ImageShapePrinter();

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ImageSavingCallback.html", options);
        }

        /// <summary>
        /// Prints the properties of each image as the saving process saves it to an image file in the local file system
        /// during the exporting of a document to HTML.
        /// </summary>
        private class ImageShapePrinter : IImageSavingCallback
        {
            void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
            {
                args.KeepImageStreamOpen = false;
                Assert.That(args.IsImageAvailable, Is.True);

                Console.WriteLine($"{args.Document.OriginalFileName.Split('\\').Last()} Image #{++mImageCount}");

                LayoutCollector layoutCollector = new LayoutCollector(args.Document);

                Console.WriteLine($"\tOn page:\t{layoutCollector.GetStartPageIndex(args.CurrentShape)}");
                Console.WriteLine($"\tDimensions:\t{args.CurrentShape.Bounds}");
                Console.WriteLine($"\tAlignment:\t{args.CurrentShape.VerticalAlignment}");
                Console.WriteLine($"\tWrap type:\t{args.CurrentShape.WrapType}");
                Console.WriteLine($"Output filename:\t{args.ImageFileName}\n");
            }

            private int mImageCount;
        }
        //ExEnd

        [TestCase(true)]
        [TestCase(false)]
        public void PrettyFormat(bool usePrettyFormat)
        {
            //ExStart
            //ExFor:SaveOptions.PrettyFormat
            //ExSummary:Shows how to enhance the readability of the raw code of a saved .html document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html) { PrettyFormat = usePrettyFormat };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.PrettyFormat.html", htmlOptions);

            // Enabling pretty format makes the raw html code more readable by adding tab stop and new line characters.
            string html = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.PrettyFormat.html");

            string newLine = Environment.NewLine;
            if (usePrettyFormat)
                Assert.That(html, Is.EqualTo($"<html>{newLine}" +
                                $"\t<head>{newLine}" +
                                    $"\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />{newLine}" +
                                    $"\t\t<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />{newLine}" +
                                    $"\t\t<meta name=\"generator\" content=\"{BuildVersionInfo.Product} {BuildVersionInfo.Version}\" />{newLine}" +
                                    $"\t\t<title>{newLine}" +
                                    $"\t\t</title>{newLine}" +
                                $"\t</head>{newLine}" +
                                $"\t<body style=\"font-family:'Times New Roman'; font-size:12pt\">{newLine}" +
                                    $"\t\t<div>{newLine}" +
                                        $"\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">{newLine}" +
                                            $"\t\t\t\t<span>Hello world!</span>{newLine}" +
                                        $"\t\t\t</p>{newLine}" +
                                        $"\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">{newLine}" +
                                            $"\t\t\t\t<span style=\"-aw-import:ignore\">&#xa0;</span>{newLine}" +
                                        $"\t\t\t</p>{newLine}" +
                                    $"\t\t</div>{newLine}" +
                                $"\t</body>{newLine}</html>"));
            else
                Assert.That(html, Is.EqualTo("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />" +
                            "<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />" +
                            $"<meta name=\"generator\" content=\"{BuildVersionInfo.Product} {BuildVersionInfo.Version}\" /><title></title></head>" +
                            "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                            "<div><p style=\"margin-top:0pt; margin-bottom:0pt\"><span>Hello world!</span></p>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"-aw-import:ignore\">&#xa0;</span></p></div></body></html>"));
            //ExEnd
        }

        [TestCase(SaveFormat.Html, "html")]
        [TestCase(SaveFormat.Mhtml, "mhtml")]
        [TestCase(SaveFormat.Epub, "epub")]
        //ExStart
        //ExFor:SaveOptions.ProgressCallback
        //ExFor:IDocumentSavingCallback
        //ExFor:IDocumentSavingCallback.Notify(DocumentSavingArgs)
        //ExFor:DocumentSavingArgs.EstimatedProgress
        //ExFor:DocumentSavingArgs
        //ExSummary:Shows how to manage a document while saving to html.
        public void ProgressCallback(SaveFormat saveFormat, string ext)
        {
            Document doc = new Document(MyDir + "Big document.docx");

            // Following formats are supported: Html, Mhtml, Epub.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(saveFormat)
            {
                ProgressCallback = new SavingProgressCallback()
            };

            var exception = Assert.Throws<OperationCanceledException>(() =>
                doc.Save(ArtifactsDir + $"HtmlSaveOptions.ProgressCallback.{ext}", saveOptions));
            Assert.That(exception?.Message.Contains("EstimatedProgress"), Is.True);
        }

        /// <summary>
        /// Saving progress callback. Cancel a document saving after the "MaxDuration" seconds.
        /// </summary>
        public class SavingProgressCallback : IDocumentSavingCallback
        {
            /// <summary>
            /// Ctr.
            /// </summary>
            public SavingProgressCallback()
            {
                mSavingStartedAt = DateTime.Now;
            }

            /// <summary>
            /// Callback method which called during document saving.
            /// </summary>
            /// <param name="args">Saving arguments.</param>
            public void Notify(DocumentSavingArgs args)
            {
                DateTime canceledAt = DateTime.Now;
                double ellapsedSeconds = (canceledAt - mSavingStartedAt).TotalSeconds;
                if (ellapsedSeconds > MaxDuration)
                    throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {canceledAt}");
            }

            /// <summary>
            /// Date and time when document saving is started.
            /// </summary>
            private readonly DateTime mSavingStartedAt;

            /// <summary>
            /// Maximum allowed duration in sec.
            /// </summary>
            private const double MaxDuration = 0.1d;
        }
        //ExEnd

        [TestCase(SaveFormat.Mobi)]
        [TestCase(SaveFormat.Azw3)]
        public void MobiAzw3DefaultEncoding(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.SaveFormat = saveFormat;
            saveOptions.Encoding = Encoding.ASCII;

            string outputFileName = $"{ArtifactsDir}HtmlSaveOptions.MobiDefaultEncoding{FileFormatUtil.SaveFormatToExtension(saveFormat)}";
            doc.Save(outputFileName);

            Encoding encoding = TestUtil.GetEncoding(outputFileName);
            Assert.That(encoding, Is.Not.EqualTo(Encoding.ASCII));
            Assert.That(encoding, Is.EqualTo(Encoding.UTF8));
        }

        [Test]
        public void HtmlReplaceBackslashWithYenSign()
        {
            //ExStart:HtmlReplaceBackslashWithYenSign
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:HtmlSaveOptions.ReplaceBackslashWithYenSign
            //ExSummary:Shows how to replace backslash characters with yen signs (Html).
            Document doc = new Document(MyDir + "Korean backslash symbol.docx");

            // By default, Aspose.Words mimics MS Word's behavior and doesn't replace backslash characters with yen signs in
            // generated HTML documents. However, previous versions of Aspose.Words performed such replacements in certain
            // scenarios. This flag enables backward compatibility with previous versions of Aspose.Words.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ReplaceBackslashWithYenSign = true;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ReplaceBackslashWithYenSign.html", saveOptions);
            //ExEnd:HtmlReplaceBackslashWithYenSign
        }

        [Test]
        public void RemoveJavaScriptFromLinks()
        {
            //ExStart:HtmlRemoveJavaScriptFromLinks
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:HtmlFixedSaveOptions.RemoveJavaScriptFromLinks
            //ExSummary:Shows how to remove JavaScript from the links.
            Document doc = new Document(MyDir + "JavaScript in HREF.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.RemoveJavaScriptFromLinks = true;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.RemoveJavaScriptFromLinks.html", saveOptions);
            //ExEnd:HtmlRemoveJavaScriptFromLinks
        }
    }
}
