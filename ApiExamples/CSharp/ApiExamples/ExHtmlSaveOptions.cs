// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fonts;
using Aspose.Words.Lists;
using NUnit.Framework;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using HtmlVersion = Aspose.Words.Saving.HtmlVersion;

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
                    Assert.That(dirFiles, Is.Empty);
                    return;

                case SaveFormat.Epub:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.AllDirectories);
                    Assert.That(dirFiles, Is.Empty);
                    return;

                case SaveFormat.Mhtml:

                    dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.AllDirectories);
                    Assert.That(dirFiles, Is.Empty);
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

        [TestCase(Aspose.Words.Saving.HtmlVersion.Html5)]
        [TestCase(Aspose.Words.Saving.HtmlVersion.Xhtml)]
        public void Html5Support(HtmlVersion htmlVersion)
        {
            Document doc = new Document(MyDir + "Document.doc");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                HtmlVersion = htmlVersion
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.Html5Support.html", saveOptions);
        }

#if !(__MOBILE__ || MAC)
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

#if !(__MOBILE__ || MAC)
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

#if !(__MOBILE__ || MAC)
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
        [Ignore("Bug")]
        public void ResolveFontNames()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ResolveFontNames
            //ExSummary:Shows how to resolve all font names before writing them to HTML.
            Document document = new Document(MyDir + "HtmlSaveOptions.ResolveFontNames.docx");

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

        [Test]
        public void HeadingLevels()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
            //ExFor:HtmlSaveOptions.DocumentSplitHeadingLevel
            //ExSummary:Shows how to split a document into several html documents by heading levels.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert headings of levels 1 - 3
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

            // Create a HtmlSaveOptions object and set the DocumentSplitHeadingLevel to 2
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
            options.DocumentSplitHeadingLevel = 2;

            // Instead of one output html, the document will be split up into 4 parts, on heading levels 1 and 2
            doc.Save(ArtifactsDir + "HeadingLevels.html", options);
            //ExEnd
        }

        [Test]
        public void NegativeIndent()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.AllowNegativeIndent
            //ExFor:HtmlSaveOptions.TableWidthOutputMode
            //ExSummary:Shows how to preserve negative indents in the output .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table and give it a negative value for its indent, effectively pushing it out of the left page boundary
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();
            table.LeftIndent = -36;
            table.PreferredWidth = PreferredWidth.FromPoints(144);

            // When saving to .html, this indent will only be preserved if we set this flag
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.AllowNegativeIndent = true;
            options.TableWidthOutputMode = HtmlElementSizeOutputMode.RelativeOnly;

            // The first cell with "Cell 1" will not be visible in the output 
            doc.Save(ArtifactsDir + "AllowNegativeIndent.html", options);
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
            //ExSummary:Shows how to set folders and folder aliases for externally saved resources when saving to html.
            Document doc = new Document(MyDir + "Rendering.doc");

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

            doc.Save(ArtifactsDir + "FolderAliasf.html", options);
            //ExEnd
        }

        [Test]
        public void HtmlVersion()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportXhtmlTransitional
            //ExFor:HtmlSaveOptions.HtmlVersion
            //ExSummary:Shows how to set a saved .html document to a specific version.
            Document doc = new Document(MyDir + "Rendering.doc");

            // Save the document to a .html file of the XHTML 1.0 Transitional standard
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                HtmlVersion = Aspose.Words.Saving.HtmlVersion.Xhtml,
                ExportXhtmlTransitional = true,
                PrettyFormat = true
            };

            // The DOCTYPE declaration at the top of this document will indicate the html version we chose
            doc.Save(ArtifactsDir + "HtmlVersion.html", options);
            //ExEnd
        }

        [Test]
        public void EpubHeadings()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.EpubNavigationMapLevel
            //ExSummary:Shows the relationship between heading levels and the Epub navigation panel.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert headings of levels 1 - 3
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

            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Epub)
            {
                EpubNavigationMapLevel = 2,
                HtmlVersion = Aspose.Words.Saving.HtmlVersion.Xhtml
            };

            doc.Save(ArtifactsDir + "EpubHeadings.epub", options);
            //ExEnd
        }

        [Test]
        public void ContentIdUrls()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportCidUrlsForMhtmlResources
            //ExSummary:Shows how to enable content IDs for output MHTML documents.
            Document doc = new Document(MyDir + "Rendering.doc");

            // Setting this flag will replace "Content-Location" tags with "Content-ID" tags for each resource from the input document
            // The file names that were next to each "Content-Location" tag are re-purposed as content IDs
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportCidUrlsForMhtmlResources = true,
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "ContentIdUrls.mht", options);
            //ExEnd
        }

        [Test]
        public void DropDownFormField()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportDropDownFormFieldAsText
            //ExSummary:Shows how to get drop down combo box form fields to blend in with paragraph text when saving to html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a combo box with the value "Two" selected
            builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 1);
            
            // When converting to .html, drop down combo boxes will be converted to select/option tags to preserve their functionality
            // If we want to freeze a combo box at its current selected value and convert it into plain text, we can do so with this flag
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportDropDownFormFieldAsText = true
            };

            doc.Save(ArtifactsDir + "DropDownFormField.html", options);
            //ExEnd
        }

        [Test]
        public void ExportBase64()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportFontsAsBase64
            //ExFor:HtmlSaveOptions.ExportImagesAsBase64
            //ExSummary:Shows how to save a .html document with resources embedded inside it.
            Document doc = new Document(MyDir + "Rendering.doc");

            // By default, when converting a document with images to .html, resources such as images will be linked to in external files
            // We can set these flags to embed resources inside the output .html instead, cutting down on the amount of files created during the conversion
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportFontsAsBase64 = true,
                ExportImagesAsBase64 = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "ExportBase64.html", options);
            //ExEnd
        }

        [Test]
        public void ExportLanguageInformation()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportLanguageInformation
            //ExSummary:Shows how to preserve language information when saving to .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use the builder to write text in more than one language
            builder.Font.LocaleId = 2057; // en-GB
            builder.Writeln("Hello world!");

            builder.Font.LocaleId = 1049; // ru-RU
            builder.Write("Привет, мир!");

            // Normally, when saving a document with more than one proofing language to .html,
            // only the text content is preserved with no traces of any other languages
            // Saving with a HtmlSaveOptions object with this flag set will add "lang" attributes to spans 
            // in places where other proofing languages were used 
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportLanguageInformation = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "ExportLanguageInformation.html", options);
            //ExEnd
        }

        [Test]
        public void List()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportListLabels
            //ExSummary:Shows how to export an indented list to .html as plain text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use the builder to insert a list
            Aspose.Words.Lists.List list = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = list;
            
            builder.Writeln("List item 1.");
            builder.ListFormat.ListIndent();
            builder.Writeln("List item 2.");
            builder.ListFormat.ListIndent();
            builder.Write("List item 3.");

            // When we save this to .html, normally our list will be represented by <li> tags
            // We can set this flag to have lists as plain text instead
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportListLabels = ExportListLabels.AsInlineText,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "List.html", options);
            //ExEnd
        }

        [Test]
        public void ExportPageMargins()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportPageMargins
            //ExSummary:Shows how to show out-of-bounds objects in output .html documents.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a builder to insert a shape with no wrapping
            Shape shape = builder.InsertShape(ShapeType.Cube, 200, 200);

            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            shape.WrapType = WrapType.None;

            // Negative values for shape position may cause the shape to go out of page bounds
            // If we export this to .html, the shape will be truncated
            shape.Left = -150;

            // We can avoid that and have the entire shape be visible by setting this flag
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportPageMargins = true
            };

            doc.Save(ArtifactsDir + "ExportPageMargins.html", options);
            //ExEnd
        }

        [Test]
        [Category("SkipTearDown")]
        public void ExportPageSetup()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportPageSetup
            //ExSummary:Shows how to preserve section structure/page setup information when saving to html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a DocumentBuilder to insert two sections with text
            builder.Writeln("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2");

            // Change dimensions and paper size of first section
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.TopMargin = 36.0;
            pageSetup.BottomMargin = 36.0;
            pageSetup.PaperSize = PaperSize.A5;

            // Section structure and pagination are normally lost when when converting to .html
            // We can create an HtmlSaveOptions object with the ExportPageSetup flag set to true
            // to preserve the section structure in <div> tags and page dimensions in the output document's CSS
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportPageSetup = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "ExportPageSetup.html", options);
            //ExEnd
        }

        [Test]
        public void RelativeFontSize()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportRelativeFontSize
            //ExSummary:Shows how to use relative font sizes when saving to .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a builder to write some text in various sizes
            builder.Writeln("Default font size, ");
            builder.Font.Size = 24.0;
            builder.Writeln("2x default font size,");
            builder.Font.Size = 96;
            builder.Write("8x default font size");

            // We can save font sizes as ratios of the default size, which will be 12 in this case
            // If we use an input .html, this size can be set with the AbsSize {font-size:12pt} tag
            // The ExportRelativeFontSize will enable this feature
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportRelativeFontSize = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "RelativeFontSize.html", options);
            //ExEnd
        }
    }
}