// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
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
using HtmlVersion = Aspose.Words.Saving.HtmlVersion;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportPageMargins(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "TextBoxes.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                SaveFormat = saveFormat,
                ExportPageMargins = true
            };

            doc.Save(ArtifactsDir +"HtmlSaveOptions.ExportPageMargins" + FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);
        }

        [TestCase(SaveFormat.Html, HtmlOfficeMathOutputMode.Image, Category = "SkipMono")]
        [TestCase(SaveFormat.Mhtml, HtmlOfficeMathOutputMode.MathML, Category = "SkipMono")]
        [TestCase(SaveFormat.Epub, HtmlOfficeMathOutputMode.Text, Category = "SkipMono")]
        public void ExportOfficeMath(SaveFormat saveFormat, HtmlOfficeMathOutputMode outputMode)
        {
            Document doc = new Document(MyDir + "Office math.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.OfficeMathOutputMode = outputMode;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportOfficeMath" + FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);
        }

        [TestCase(SaveFormat.Html, true, Description = "TextBox as svg (html)")]
        [TestCase(SaveFormat.Epub, true, Description = "TextBox as svg (epub)")]
        [TestCase(SaveFormat.Mhtml, false, Description = "TextBox as img (mhtml)")]
        public void ExportTextBoxAsSvg(SaveFormat saveFormat, bool isTextBoxAsSvg)
        {
            string[] dirFiles;

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textbox = builder.InsertShape(ShapeType.TextBox, 300, 100);
            builder.MoveTo(textbox.FirstParagraph);
            builder.Write("Hello world!");

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
                // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
                // otherwise HTML <p> tag is used. This is also the default value
                // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation
                // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible
                ExportListLabels = howExportListLabels
            };

            doc.Save(ArtifactsDir + $"HtmlSaveOptions.ControlListLabelsExport.html", saveOptions);
        }

        [TestCase(true)]
        [TestCase(false)]
        public void ExportUrlForLinkedImage(bool export)
        {
            Document doc = new Document(MyDir + "Linked image.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportOriginalUrlForLinkedImages = export };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

            string[] dirFiles = Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportUrlForLinkedImage.001.png", SearchOption.AllDirectories);

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
            //Assert that default value is true for HTML and false for MHTML and EPUB
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

            string[] imageFiles = Directory.GetFiles(ArtifactsDir + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.png", SearchOption.AllDirectories);
            Assert.AreEqual(8, imageFiles.Length);

            string[] fontFiles = Directory.GetFiles(ArtifactsDir + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.ttf", SearchOption.AllDirectories);
            Assert.AreEqual(10, fontFiles.Length);

            string[] cssFiles = Directory.GetFiles(ArtifactsDir + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.css", SearchOption.AllDirectories);
            Assert.AreEqual(1, cssFiles.Length);

            DocumentHelper.FindTextInFile(ArtifactsDir + "HtmlSaveOptions.ExternalResourceSavingConfig.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExternalResourceSavingConfig.css\"");
        }

        [Test]
        public void ConvertFontsAsBase64()
        {
            Document doc = new Document(MyDir + "TextBoxes.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ResourceFolder = "Resources";
            saveOptions.ExportFontResources = true;
            saveOptions.ExportFontsAsBase64 = true;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ConvertFontsAsBase64.html", saveOptions);
		}

        [TestCase(Aspose.Words.Saving.HtmlVersion.Html5)]
        [TestCase(Aspose.Words.Saving.HtmlVersion.Xhtml)]
        public void Html5Support(HtmlVersion htmlVersion)
        {
            Document doc = new Document(MyDir + "Document.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                HtmlVersion = htmlVersion
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.Html5Support.html", saveOptions);
        }

#if NET462 || NETCOREAPP2_1 || JAVA
        [TestCase(false)]
        [TestCase(true)]
        public void ExportFonts(bool exportAsBase64)
        {
            Document doc = new Document(MyDir + "Document.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,
                ExportFontsAsBase64 = exportAsBase64
            };

            switch (exportAsBase64)
            {
                case false:

                    doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportFonts.False.html", saveOptions);
                    Assert.IsNotEmpty(Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportFonts.False.times.ttf",
                        SearchOption.AllDirectories));
                    break;

                case true:

                    doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportFonts.True.html", saveOptions);
                    Assert.IsEmpty(Directory.GetFiles(ArtifactsDir, "HtmlSaveOptions.ExportFonts.True.times.ttf",
                        SearchOption.AllDirectories));
                    break;
            }
        }

        [Test]
        public void ResourceFolderPriority()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.ExportFontResources = true;
            saveOptions.ResourceFolder = ArtifactsDir + "Resources";
            saveOptions.ResourceFolderAlias = "http://example.com/resources";

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ResourceFolderPriority.html", saveOptions);

            string[] a = Directory.GetFiles(ArtifactsDir + "Resources", "HtmlSaveOptions.ResourceFolderPriority.001.jpeg",
                SearchOption.AllDirectories);
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
            //ExSummary:Shows how to specifies a prefix which is added to all CSS class names.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.Embedded,
                CssClassNamePrefix = "myprefix-"
            };

            // The prefix will be found before CSS element names in the embedded stylesheet
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

        [Test]
        [Ignore("Bug")]
        public void ResolveFontNames()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ResolveFontNames
            //ExSummary:Shows how to resolve all font names before writing them to HTML.
            Document document = new Document(MyDir + "Missing font.docx");

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
                // By default this option is set to 'False' and Aspose.Words writes font names as specified in the source document
                ResolveFontNames = true 
            };

            document.Save(ArtifactsDir + "HtmlSaveOptions.ResolveFontNames.html", saveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlSaveOptions.ResolveFontNames.html");

            Assert.True(Regex.Match(outDocContents, "<span style=\"font-family:Arial\">").Success);
            //ExEnd
        }

        [Test]
        public void HeadingLevels()
        {
            //ExStart
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

            // Create a HtmlSaveOptions object and set the split criteria to "HeadingParagraph", meaning that the document 
            // will be split into parts at the beginning of every paragraph of a "Heading" style, and each part will be saved as a separate document
            // Also, we will set the DocumentSplitHeadingLevel to 2, which will split the document only at headings that have levels from 1 to 2
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 2
            };
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.HeadingLevels.html", options);
            //ExEnd
        }

        [Test]
        public void NegativeIndent()
        {
            //ExStart
            //ExFor:HtmlElementSizeOutputMode
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
            doc.Save(ArtifactsDir + "HtmlSaveOptions.NegativeIndent.html", options);
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

        [Test]
        public void HtmlVersion()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
            //ExFor:HtmlSaveOptions.ExportXhtmlTransitional
            //ExFor:HtmlSaveOptions.HtmlVersion
            //ExFor:HtmlVersion
            //ExSummary:Shows how to set a saved .html document to a specific version.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Save the document to a .html file of the XHTML 1.0 Transitional standard
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                HtmlVersion = Aspose.Words.Saving.HtmlVersion.Xhtml,
                ExportXhtmlTransitional = true,
                PrettyFormat = true
            };

            // The DOCTYPE declaration at the top of this document will indicate the html version we chose
            doc.Save(ArtifactsDir + "HtmlSaveOptions.HtmlVersion.html", options);
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

            // Epub readers normally treat paragraphs with "Heading" styles as anchors for a table of contents-style navigation pane
            // We set a maximum heading level above which headings won't be registered by the reader as navigation points with
            // a HtmlSaveOptions object and its EpubNavigationLevel attribute
            // Our document has headings of levels 1 to 3,
            // but our output epub will only place level 1 and 2 headings in the table of contents
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Epub);
            options.EpubNavigationMapLevel = 2;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.EpubHeadings.epub", options);
            //ExEnd
        }

        [Test]
        public void ContentIdUrls()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportCidUrlsForMhtmlResources
            //ExSummary:Shows how to enable content IDs for output MHTML documents.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Setting this flag will replace "Content-Location" tags with "Content-ID" tags for each resource from the input document
            // The file names that were next to each "Content-Location" tag are re-purposed as content IDs
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportCidUrlsForMhtmlResources = true,
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ContentIdUrls.mht", options);
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
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportDropDownFormFieldAsText = true;    

            doc.Save(ArtifactsDir + "HtmlSaveOptions.DropDownFormField.html", options);
            //ExEnd
        }

        [Test]
        public void ExportBase64()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportFontsAsBase64
            //ExFor:HtmlSaveOptions.ExportImagesAsBase64
            //ExSummary:Shows how to save a .html document with resources embedded inside it.
            Document doc = new Document(MyDir + "Rendering.docx");

            // By default, when converting a document with images to .html, resources such as images will be linked to in external files
            // We can set these flags to embed resources inside the output .html instead, cutting down on the amount of files created during the conversion
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportFontsAsBase64 = true,
                ExportImagesAsBase64 = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportBase64.html", options);
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
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportLanguageInformation = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportLanguageInformation.html", options);
            //ExEnd
        }

        [Test]
        public void List()
        {
            //ExStart
            //ExFor:ExportListLabels
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
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportListLabels = ExportListLabels.AsInlineText,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.List.html", options);
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
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportPageMargins = true;
        
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportPageMargins.html", options);
            //ExEnd
        }

        [Test]
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
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportPageSetup = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportPageSetup.html", options);
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
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportRelativeFontSize = true,
                PrettyFormat = true
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.RelativeFontSize.html", options);
            //ExEnd
        }

        [Test]
        public void ExportTextBox()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportTextBoxAsSvg
            //ExSummary:Shows how to export text boxes as scalable vector graphics.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a DocumentBuilder to insert a text box and give it some text content
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100.0, 60.0);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("My text box");

            // Normally, all shapes such as the text box we placed are exported to .html as external images linked by the .html document
            // We can save with an HtmlSaveOptions object with the ExportTextBoxAsSvg set to true to save text boxes as <svg> tags,
            // which will cause no linked images to be saved and will make the inner text selectable
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportTextBoxAsSvg = true;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTextBox.html", options);
            //ExEnd
        }

        [Test]
        public void RoundTripInformation()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportRoundtripInformation
            //ExSummary:Shows how to preserve hidden elements when converting to .html.
            Document doc = new Document(MyDir + "Rendering.docx");

            // When converting a document to .html, some elements such as hidden bookmarks, original shape positions,
            // or footnotes will be either removed or converted to plain text and effectively be lost
            // Saving with a HtmlSaveOptions object with ExportRoundtripInformation set to true will preserve these elements
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportRoundtripInformation = true,
                PrettyFormat = true
            };

            // These elements will have tags that will start with "-aw", such as "-aw-import" or "-aw-left-pos"
            doc.Save(ArtifactsDir + "HtmlSaveOptions.RoundTripInformation.html", options);
            //ExEnd
        }

        [Test]
        public void ExportTocPageNumbers()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportTocPageNumbers
            //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents
            FieldToc fieldToc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

            // Populate the document with paragraphs of a "Heading" style that the table of contents will pick up
            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 1"];
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Entry 1");
            builder.Writeln("Entry 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Entry 3");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Entry 4");

            // Our headings span several pages, and those page numbers will be displayed by the TOC at the top of the document
            fieldToc.UpdatePageNumbers();
            doc.UpdateFields();

            // These page numbers are normally omitted since .html has no pagination, but we can still have them displayed
            // if we save with a HtmlSaveOptions object with the ExportTocPageNumbers set to true 
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportTocPageNumbers = true;
            
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ExportTocPageNumbers.html", options);
            //ExEnd
        }

        [Test]
        public void FontSubsetting()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.FontResourcesSubsettingSizeThreshold
            //ExSummary:Shows how to work with font subsetting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a DocumentBuilder to insert text with several fonts
            builder.Font.Name = "Arial";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Times New Roman";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Courier New";
            builder.Writeln("Hello world!");

            // When saving to .html, font subsetting fully applies by default, meaning that when we export fonts with our file,
            // the symbols not used by our document are not represented by the exported fonts, which cuts down file size dramatically
            // Font files of a file size larger than FontResourcesSubsettingSizeThreshold get subsetted, so a value of 0 will apply default full subsetting
            // Setting the value to something large will fully suppress subsetting, saving some very large font files that cover every glyph
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportFontResources = true,
                FontResourcesSubsettingSizeThreshold = int.MaxValue
            };

            doc.Save(ArtifactsDir + "HtmlSaveOptions.FontSubsetting.html", options);
            //ExEnd
        }

        [Test]
        public void MetafileFormat()
        {
            //ExStart
            //ExFor:HtmlMetafileFormat
            //ExFor:HtmlSaveOptions.MetafileFormat
            //ExSummary:Shows how to set a meta file in a different format.
            // Create a document from an html string
            string html = 
                @"<html>
                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>
                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>
                    </svg>
                </html>";

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)));

            // This document contains a <svg> element in the form of text,
            // which by default will be saved as a linked external .png when we save the document as html
            // We can save with a HtmlSaveOptions object with this flag set to preserve the <svg> tag
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.MetafileFormat = HtmlMetafileFormat.Svg;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.MetafileFormat.html", options);
            //ExEnd
        }

        [Test]
        public void OfficeMathOutputMode()
        {
            //ExStart
            //ExFor:HtmlOfficeMathOutputMode
            //ExFor:HtmlSaveOptions.OfficeMathOutputMode
            //ExSummary:Shows how to control the way how OfficeMath objects are exported to .html.
            // Open a document that contains OfficeMath objects
            Document doc = new Document(MyDir + "Office math.docx");

            // Create a HtmlSaveOptions object and configure it to export OfficeMath objects as images
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.OfficeMathOutputMode = HtmlOfficeMathOutputMode.Image;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.OfficeMathOutputMode.html", options);
            //ExEnd
        }

        [Test]
        public void ScaleImageToShapeSize()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ScaleImageToShapeSize
            //ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
            // Open a document which contains shapes with images
            Document doc = new Document(MyDir + "Rendering.docx");

            // By default, images inside shapes get scaled to the size of their shapes while the document gets 
            // converted to .html, reducing image file size
            // We can save the document with a HtmlSaveOptions with ScaleImageToShapeSize set to false to prevent the scaling
            // and preserve the full quality and file size of the linked images
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ScaleImageToShapeSize = false;

            doc.Save(ArtifactsDir + "HtmlSaveOptions.ScaleImageToShapeSize.html", options);
            //ExEnd
        }

        //ExStart
        //ExFor:ImageSavingArgs.CurrentShape
        //ExFor:ImageSavingArgs.Document
        //ExFor:ImageSavingArgs.ImageStream
        //ExFor:ImageSavingArgs.IsImageAvailable
        //ExFor:ImageSavingArgs.KeepImageStreamOpen
        //ExSummary:Shows how to involve an image saving callback in an .html conversion process.
        [Test] //ExSkip
        public void ImageSavingCallback()
        {
            // Open a document which contains shapes with images
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create a HtmlSaveOptions object with a custom image saving callback that will print image information
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ImageSavingCallback = new ImageShapePrinter();
           
            doc.Save(ArtifactsDir + "HtmlSaveOptions.ImageSavingCallback.html", options);
        }

        /// <summary>
        /// Prints information of all images that are about to be saved from within a document to image files
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
                Console.WriteLine($"\tDimensions:\t{args.CurrentShape.Bounds.ToString()}");
                Console.WriteLine($"\tAlignment:\t{args.CurrentShape.VerticalAlignment}");
                Console.WriteLine($"\tWrap type:\t{args.CurrentShape.WrapType}");
                Console.WriteLine($"Output filename:\t{args.ImageFileName}\n");
            }

            private int mImageCount;
        }
        //ExEnd
    }
}