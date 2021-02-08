// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
using NUnit.Framework;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
#if NETCOREAPP2_1 || __MOBILE__
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
        public void ExportTextBoxAsSvgEpub(SaveFormat saveFormat, bool isTextBoxAsSvg)
        {
            string[] dirFiles;

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textbox = builder.InsertShape(ShapeType.TextBox, 300, 100);
            builder.MoveTo(textbox.FirstParagraph);
            builder.Write("Hello world!");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(saveFormat);
            saveOptions.ExportTextBoxAsSvg = isTextBoxAsSvg;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTextBoxAsSvgEpub" + FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);

            switch (saveFormat)
            {
                case SaveFormat.Html:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                        SearchOption.AllDirectories);
                    Assert.That(dirFiles, Is.Empty);
                    return;

                case SaveFormat.Epub:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                        SearchOption.AllDirectories);
                    Assert.That(dirFiles, Is.Empty);
                    return;

                case SaveFormat.Mhtml:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                        SearchOption.AllDirectories);
                    Assert.That(dirFiles, Is.Empty);
                    return;
            }
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
            Assert.AreEqual(true, saveOptions.ExportRoundtripInformation);

            saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
            Assert.AreEqual(false, saveOptions.ExportRoundtripInformation);

            saveOptions = new HtmlSaveOptions(SaveFormat.Epub);
            Assert.AreEqual(false, saveOptions.ExportRoundtripInformation);
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
            Assert.AreEqual(8, imageFiles.Length);

            string[] fontFiles = Directory.GetFiles(ArtifactsDir + "Resources/",
                "HtmlSaveOptions.ExternalResourceSavingConfig*.ttf", SearchOption.AllDirectories);
            Assert.AreEqual(10, fontFiles.Length);

            string[] cssFiles = Directory.GetFiles(ArtifactsDir + "Resources/",
                "HtmlSaveOptions.ExternalResourceSavingConfig*.css", SearchOption.AllDirectories);
            Assert.AreEqual(1, cssFiles.Length);

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

#if NET462 || NETCOREAPP2_1 || JAVA
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

                    Assert.IsNotEmpty(Directory.GetFiles(fontsFolder, "HtmlSaveOptions.ExportFonts.False.times.ttf",
                        SearchOption.AllDirectories));

                    Directory.Delete(fontsFolder, true);
                    break;

                case true:

                    doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportFonts.True.html", saveOptions);
                    Assert.False(Directory.Exists(fontsFolder));
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

            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.001.png", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.002.png", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.arial.ttf", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.css", SearchOption.AllDirectories));
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

            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Images",
                "HtmlSaveOptions.ResourceFolderLowPriority.001.png", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Images", "HtmlSaveOptions.ResourceFolderLowPriority.002.png",
                SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Fonts",
                "HtmlSaveOptions.ResourceFolderLowPriority.arial.ttf", SearchOption.AllDirectories));
            Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderLowPriority.css",
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

            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.SvgMetafileFormat.html",
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

            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.PngMetafileFormat.html",
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

            builder.Document.Save(ArtifactsDir + "HtmlSaveOptions.EmfOrWmfMetafileFormat.html",
                new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.EmfOrWmf });
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

            Assert.True(outDocContents.Contains("<p class=\"myprefix-Header\">"));
            Assert.True(outDocContents.Contains("<p class=\"myprefix-Footer\">"));

            outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.CssClassNamePrefix.css");

            Assert.True(outDocContents.Contains(".myprefix-Footer { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\r\n" +
                                                ".myprefix-Header { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\r\n"));
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
            Assert.NotNull(doc.FontInfos["28 Days Later"]);

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

            Assert.True(resolveFontNames
                ? Regex.Match(outDocContents, "<span style=\"font-family:Arial\">").Success
                : Regex.Match(outDocContents, "<span style=\"font-family:\'28 Days Later\'\">").Success);
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

            Assert.AreEqual("Heading #1", doc.GetText().Trim());

            doc = new Document(ArtifactsDir + "HtmlSaveOptions.HeadingLevels-01.html");

            Assert.AreEqual("Heading #2\r" +
                            "Heading #3", doc.GetText().Trim());

            doc = new Document(ArtifactsDir + "HtmlSaveOptions.HeadingLevels-02.html");

            Assert.AreEqual("Heading #4", doc.GetText().Trim());

            doc = new Document(ArtifactsDir + "HtmlSaveOptions.HeadingLevels-03.html");

            Assert.AreEqual("Heading #5\r" +
                            "Heading #6", doc.GetText().Trim());
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
                Assert.True(outDocContents.Contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:-41.65pt; border:0.75pt solid #000000; -aw-border:0.5pt single; border-collapse:collapse\">"));
                Assert.True(outDocContents.Contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; border-collapse:collapse\">"));
            }
            else
            {
                Assert.True(outDocContents.Contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"border:0.75pt solid #000000; -aw-border:0.5pt single; border-collapse:collapse\">"));
                Assert.True(outDocContents.Contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; border-collapse:collapse\">"));
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
            {
                Console.WriteLine(fontFilename);
            }

            Assert.AreEqual(10, Array.FindAll(Directory.GetFiles(ArtifactsDir), s => s.EndsWith(".ttf")).Length); //ExSkip
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
                Assert.True(args.Document.OriginalFileName.EndsWith("Rendering.docx"));

                Assert.True(args.IsExportNeeded);
                Assert.True(args.IsSubsettingNeeded);

                // There are two ways of saving an exported font.
                // 1 -  Save it to a local file system location:
                args.FontFileName = args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last();

                // 2 -  Save it to a stream:
                args.FontStream =
                    new FileStream(ArtifactsDir + args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last(), FileMode.Create);
                Assert.False(args.KeepFontStreamOpen);
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
                    Assert.True(outDocContents.Contains("<a id=\"_Toc76372689\"></a>"));
                    Assert.True(outDocContents.Contains("<a id=\"_Toc76372689\"></a>"));
                    Assert.True(outDocContents.Contains("<table style=\"border-collapse:collapse\">"));
                    break;
                case HtmlVersion.Xhtml:
                    Assert.True(outDocContents.Contains("<a name=\"_Toc76372689\"></a>"));
                    Assert.True(outDocContents.Contains("<ul type=\"disc\" style=\"margin:0pt; padding-left:0pt\">"));
                    Assert.True(outDocContents.Contains("<table cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse\">"));
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

            if (showDoctypeDeclaration)
                Assert.True(outDocContents.Contains(
                    "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>\r\n" +
                    "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" +
                    "<html xmlns=\"http://www.w3.org/1999/xhtml\">"));
            else
                Assert.True(outDocContents.Contains("<html>"));
            //ExEnd
        }

        [Test]
        public void EpubHeadings()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.EpubNavigationMapLevel
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
            // We can use the "EpubNavigationMapLevel" property to set a maximum heading level. 
            // The Epub reader will not add headings with a level above the one we specify to the contents table.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Epub);
            options.EpubNavigationMapLevel = 2;

            // Our document has six headings, two of which are above level 2.
            // The table of contents for this document will have four entries.
            doc.Save(ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", options);
            //ExEnd

            TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #1</text></navLabel>", 
                ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
            TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #2</text></navLabel>", 
                ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
            TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #4</text></navLabel>", 
                ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
            TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #5</text></navLabel>", 
                ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");

            Assert.Throws<AssertionException>(() =>
            {
                TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #3</text></navLabel>", 
                    ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
            });

            Assert.Throws<AssertionException>(() =>
            {
                TestUtil.DocPackageFileContainsString("<navLabel><text>Heading #6</text></navLabel>", 
                    ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
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
                Assert.True(outDocContents.Contains("Content-ID: <document.html>"));
                Assert.True(outDocContents.Contains("<link href=3D\"cid:styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
                Assert.True(outDocContents.Contains("@font-face { font-family:'Arial Black'; src:url('cid:ariblk.ttf') }"));
                Assert.True(outDocContents.Contains("<img src=3D\"cid:image.003.jpeg\" width=3D\"351\" height=3D\"180\" alt=3D\"\" />"));
            }
            else
            {
                Assert.True(outDocContents.Contains("Content-Location: document.html"));
                Assert.True(outDocContents.Contains("<link href=3D\"styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
                Assert.True(outDocContents.Contains("@font-face { font-family:'Arial Black'; src:url('ariblk.ttf') }"));
                Assert.True(outDocContents.Contains("<img src=3D\"image.003.jpeg\" width=3D\"351\" height=3D\"180\" alt=3D\"\" />"));
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
                Assert.True(outDocContents.Contains(
                    "<span>Two</span>"));
            else
                Assert.True(outDocContents.Contains(
                    "<select name=\"MyComboBox\">" +
                        "<option>One</option>" +
                        "<option selected=\"selected\">Two</option>" +
                        "<option>Three</option>" +
                    "</select>"));
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportImagesAsBase64(bool exportItemsAsBase64)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportFontsAsBase64
            //ExFor:HtmlSaveOptions.ExportImagesAsBase64
            //ExSummary:Shows how to save a .html document with images embedded inside it.
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = exportItemsAsBase64,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportImagesAsBase64.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportImagesAsBase64.html");

            Assert.True(exportItemsAsBase64
                ? outDocContents.Contains("<img src=\"data:image/png;base64")
                : outDocContents.Contains("<img src=\"HtmlSaveOptions.ExportImagesAsBase64.001.png\""));
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
                Assert.True(outDocContents.Contains("<span>Hello world!</span>"));
                Assert.True(outDocContents.Contains("<span lang=\"en-GB\">Hello again!</span>"));
                Assert.True(outDocContents.Contains("<span lang=\"ru-RU\">Привет, мир!</span>"));
            }
            else
            {
                Assert.True(outDocContents.Contains("<span>Hello world!</span>"));
                Assert.True(outDocContents.Contains("<span>Hello again!</span>"));
                Assert.True(outDocContents.Contains("<span>Привет, мир!</span>"));
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
            
            Aspose.Words.Lists.List list = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = list;
            
            builder.Writeln("Default numbered list item 1.");
            builder.Writeln("Default numbered list item 2.");
            builder.ListFormat.ListIndent();
            builder.Writeln("Default numbered list item 3.");
            builder.ListFormat.RemoveNumbers();

            list = doc.Lists.Add(ListTemplate.OutlineHeadingsLegal);
            builder.ListFormat.List = list;

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
                    Assert.True(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-left:72pt; margin-bottom:0pt; text-indent:-18pt; -aw-import:list-item; -aw-list-level-number:1; -aw-list-number-format:'%1.'; -aw-list-number-styles:'lowerLetter'; -aw-list-number-values:'1'; -aw-list-padding-sml:9.67pt\">" +
                            "<span style=\"-aw-import:ignore\">" +
                                "<span>a.</span>" +
                                "<span style=\"font:7pt 'Times New Roman'; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "</span>" +
                            "<span>Default numbered list item 3.</span>" +
                        "</p>"));

                    Assert.True(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; -aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; -aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" +
                            "<span style=\"-aw-import:ignore\">" +
                                "<span>2.1.1.1</span>" +
                                "<span style=\"font:7pt 'Times New Roman'; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "</span>" +
                            "<span>Outline legal heading list item 5.</span>" +
                        "</p>"));
                    break;
                case ExportListLabels.Auto:
                    Assert.True(outDocContents.Contains(
                        "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                            "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
                                "<span>Default numbered list item 3.</span>" +
                            "</li>" +
                        "</ol>"));

                    Assert.True(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; " +
                        "-aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; " +
                        "-aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" +
                            "<span style=\"-aw-import:ignore\">" +
                                "<span>2.1.1.1</span>" +
                                "<span style=\"font:7pt 'Times New Roman'; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "</span>" +
                            "<span>Outline legal heading list item 5.</span>" +
                        "</p>"));
                    break;
                case ExportListLabels.ByHtmlTags:
                    Assert.True(outDocContents.Contains(
                        "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                            "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
                                "<span>Default numbered list item 3.</span>" +
                            "</li>" +
                        "</ol>"));

                    Assert.True(outDocContents.Contains(
                        "<ol type=\"1\" class=\"awlist3\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                            "<li style=\"margin-left:7.2pt; text-indent:-43.2pt; -aw-list-padding-sml:10.2pt\">" +
                                "<span style=\"font:7pt 'Times New Roman'; -aw-import:ignore\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                                "<span>Outline legal heading list item 5.</span>" +
                            "</li>" +
                        "</ol>"));
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
                Assert.True(outDocContents.Contains("<style type=\"text/css\">div.Section1 { margin:70.85pt }</style>"));
                Assert.True(outDocContents.Contains("<div class=\"Section1\"><p style=\"margin-top:0pt; margin-left:151pt; margin-bottom:0pt\">"));
            }
            else
            {
                Assert.False(outDocContents.Contains("style type=\"text/css\">"));
                Assert.True(outDocContents.Contains("<div><p style=\"margin-top:0pt; margin-left:221.85pt; margin-bottom:0pt\">"));
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
                Assert.True(outDocContents.Contains(
                    "<style type=\"text/css\">" +
                        "@page Section1 { size:419.55pt 595.3pt; margin:36pt 70.85pt }" +
                        "@page Section2 { size:612pt 792pt; margin:70.85pt }" +
                        "div.Section1 { page:Section1 }div.Section2 { page:Section2 }" +
                    "</style>"));

                Assert.True(outDocContents.Contains(
                    "<div class=\"Section1\">" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Section 1</span>" +
                        "</p>" +
                    "</div>"));
            }
            else
            {
                Assert.False(outDocContents.Contains("style type=\"text/css\">"));

                Assert.True(outDocContents.Contains(
                    "<div>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Section 1</span>" +
                        "</p>" +
                    "</div>"));
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
                Assert.True(outDocContents.Contains(
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
                    "</body>"));
            }
            else
            {
                Assert.True(outDocContents.Contains(
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
                    "</body>"));
            }
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportTextBox(bool exportTextBoxAsSvg)
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportTextBoxAsSvg
            //ExSummary:Shows how to export text boxes as scalable vector graphics.
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
            HtmlSaveOptions options = new HtmlSaveOptions { ExportTextBoxAsSvg = exportTextBoxAsSvg };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTextBox.html", options);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ExportTextBox.html");

            if (exportTextBoxAsSvg)
            {
                Assert.True(outDocContents.Contains(
                    "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"133\" height=\"80\">"));
            }
            else
            {
                Assert.True(outDocContents.Contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                        "<img src=\"HtmlSaveOptions.ExportTextBox.001.png\" width=\"136\" height=\"83\" alt=\"\" " +
                        "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                    "</p>"));
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
                Assert.True(outDocContents.Contains("<div style=\"-aw-headerfooter-type:header-primary; clear:both\">"));
                Assert.True(outDocContents.Contains("<span style=\"-aw-import:ignore\">&#xa0;</span>"));
                
                Assert.True(outDocContents.Contains(
                    "td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                    "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top; " +
                    "-aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single; -aw-border-top:0.5pt single\">"));
                
                Assert.True(outDocContents.Contains(
                    "<li style=\"margin-left:30.2pt; padding-left:5.8pt; -aw-font-family:'Courier New'; -aw-font-weight:normal; -aw-number-format:'o'\">"));
                
                Assert.True(outDocContents.Contains(
                    "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"351\" height=\"180\" alt=\"\" " +
                    "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />"));


                Assert.True(outDocContents.Contains(
                    "<span>Page number </span>" +
                    "<span style=\"-aw-field-start:true\"></span>" +
                    "<span style=\"-aw-field-code:' PAGE   \\\\* MERGEFORMAT '\"></span>" +
                    "<span style=\"-aw-field-separator:true\"></span>" +
                    "<span>1</span>" +
                    "<span style=\"-aw-field-end:true\"></span>"));

                Assert.AreEqual(1, doc.Range.Fields.Count(f => f.Type == FieldType.FieldPage));
            }
            else
            {
                Assert.True(outDocContents.Contains("<div style=\"clear:both\">"));
                Assert.True(outDocContents.Contains("<span>&#xa0;</span>"));
                
                Assert.True(outDocContents.Contains(
                    "<td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                    "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top\">"));
                
                Assert.True(outDocContents.Contains(
                    "<li style=\"margin-left:30.2pt; padding-left:5.8pt\">"));
                
                Assert.True(outDocContents.Contains(
                    "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"351\" height=\"180\" alt=\"\" />"));

                Assert.True(outDocContents.Contains(
                    "<span>Page number 1</span>"));

                Assert.AreEqual(0, doc.Range.Fields.Count(f => f.Type == FieldType.FieldPage));
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
                Assert.True(outDocContents.Contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                    "<span>Entry 1</span>" +
                    "<span style=\"width:428.14pt; font-family:'Lucida Console'; font-size:10pt; display:inline-block; -aw-font-family:'Times New Roman'; " +
                    "-aw-tabstop-align:right; -aw-tabstop-leader:dots; -aw-tabstop-pos:469.8pt\">.......................................................................</span>" +
                    "<span>2</span>" +
                    "</p>"));
            }
            else
            {
                Assert.True(outDocContents.Contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                    "<span>Entry 1</span>" +
                    "</p>"));
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

            Assert.AreEqual(3, fontFileNames.Length);

            foreach (string filename in fontFileNames)
            {
                // By default, the .ttf files for each of our three fonts will be over 700MB.
                // Subsetting will reduce them all to under 30MB.
                FileInfo fontFileInfo = new FileInfo(filename);

                Assert.True(fontFileInfo.Length > 700000 || fontFileInfo.Length < 30000);
                Assert.True(Math.Max(fontResourcesSubsettingSizeThreshold, 30000) > new FileInfo(filename).Length);
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
            //ExSummary:Shows how to convert SVG objects to a different format when saving HTML documents.
            string html = 
                @"<html>
                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>
                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>
                    </svg>
                </html>";

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)));

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
                    Assert.True(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.MetafileFormat.001.png\" width=\"500\" height=\"40\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>"));
                    break;
                case HtmlMetafileFormat.Svg:
                    Assert.True(outDocContents.Contains(
                        "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                        "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"499\" height=\"40\">"));
                    break;
                case HtmlMetafileFormat.EmfOrWmf:
                    Assert.True(outDocContents.Contains(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.MetafileFormat.001.emf\" width=\"500\" height=\"40\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>"));
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
                    Assert.True(Regex.Match(outDocContents, 
                        "<p style=\"margin-top:0pt; margin-bottom:10pt\">" +
                            "<img src=\"HtmlSaveOptions.OfficeMathOutputMode.001.png\" width=\"159\" height=\"19\" alt=\"\" style=\"vertical-align:middle; " +
                            "-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>").Success);
                    break;
                case HtmlOfficeMathOutputMode.MathML:
                    Assert.True(Regex.Match(outDocContents, 
                        "<p style=\"margin-top:0pt; margin-bottom:10pt\">" +
                            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                                "<mi>i</mi>" +
                                "<mo>[+]</mo>" +
                                "<mi>b</mi>" +
                                "<mo>-</mo>" +
                                "<mi>c</mi>" +
                                "<mo>≥</mo>" +
                                ".*" +
                            "</math>" +
                        "</p>").Success);
                    break;
                case HtmlOfficeMathOutputMode.Text:
                    Assert.True(Regex.Match(outDocContents,
                        @"<p style=\""margin-top:0pt; margin-bottom:10pt\"">" +
                            @"<span style=\""font-family:'Cambria Math'\"">i[+]b-c≥iM[+]bM-cM </span>" +
                        "</p>").Success);
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
#if NET462 || JAVA
            Image image = Image.FromFile(ImageDir + "Transparent background logo.png");

            Assert.AreEqual(400, image.Size.Width);
            Assert.AreEqual(400, image.Size.Height);
#elif NETCOREAPP2_1
            SKBitmap image = SKBitmap.Decode(ImageDir + "Transparent background logo.png");

            Assert.AreEqual(400, image.Width);
            Assert.AreEqual(400, image.Height);
#endif

            Shape imageShape = builder.InsertImage(image);
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

            FileInfo fileInfo = new FileInfo(ArtifactsDir + "HtmlSaveOptions.ScaleImageToShapeSize.001.png");

#if NET462 || JAVA
        if (scaleImageToShapeSize)
            Assert.That(3000, Is.AtLeast(fileInfo.Length));
        else
            Assert.That(20000, Is.LessThan(fileInfo.Length));
#elif NETCOREAPP2_1
        if (scaleImageToShapeSize)
            Assert.That(10000, Is.AtLeast(fileInfo.Length));
        else
            Assert.That(30000, Is.LessThan(fileInfo.Length));
#endif
            //ExEnd
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

            Assert.IsTrue(File.Exists(ArtifactsDir + "HtmlSaveOptions.SaveHtmlWithOptions.html"));
            Assert.AreEqual(9, Directory.GetFiles(imagesDir).Length);

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
                Assert.True(args.IsImageAvailable);

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

            if (usePrettyFormat)
                Assert.AreEqual(
                    "<html>\r\n" +
                                "\t<head>\r\n" +
                                    "\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n" +
                                    "\t\t<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />\r\n" +
                                    $"\t\t<meta name=\"generator\" content=\"{BuildVersionInfo.Product} {BuildVersionInfo.Version}\" />\r\n" +
                                    "\t\t<title>\r\n" +
                                    "\t\t</title>\r\n" +
                                "\t</head>\r\n" +
                                "\t<body style=\"font-family:'Times New Roman'; font-size:12pt\">\r\n" +
                                    "\t\t<div>\r\n" +
                                        "\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" +
                                            "\t\t\t\t<span>Hello world!</span>\r\n" +
                                        "\t\t\t</p>\r\n" +
                                        "\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" +
                                            "\t\t\t\t<span style=\"-aw-import:ignore\">&#xa0;</span>\r\n" +
                                        "\t\t\t</p>\r\n" +
                                    "\t\t</div>\r\n" +
                                "\t</body>\r\n</html>", 
                    html);
            else
                Assert.AreEqual(
                    "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />" +
                            "<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />" +
                            $"<meta name=\"generator\" content=\"{BuildVersionInfo.Product} {BuildVersionInfo.Version}\" /><title></title></head>" +
                            "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                            "<div><p style=\"margin-top:0pt; margin-bottom:0pt\"><span>Hello world!</span></p>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"-aw-import:ignore\">&#xa0;</span></p></div></body></html>", 
                    html);
            //ExEnd
        }
    }
}