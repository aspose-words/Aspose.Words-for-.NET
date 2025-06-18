// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Fonts;
using Aspose.Words.Notes;
using Aspose.Words.Settings;
using Aspose.Words.Tables;
using Aspose.Words.Themes;
using NUnit.Framework;
using Font = Aspose.Words.Font;

namespace ApiExamples
{
    [TestFixture]
    public class ExFont : ApiExampleBase
    {
        [Test]
        public void CreateFormattedRun()
        {
            //ExStart
            //ExFor:Document.#ctor
            //ExFor:Font
            //ExFor:Font.Name
            //ExFor:Font.Size
            //ExFor:Font.HighlightColor
            //ExFor:Run
            //ExFor:Run.#ctor(DocumentBase,String)
            //ExFor:Story.FirstParagraph
            //ExSummary:Shows how to format a run of text using its font property.
            Document doc = new Document();
            Run run = new Run(doc, "Hello world!");

            Aspose.Words.Font font = run.Font;
            font.Name = "Courier New";
            font.Size = 36;
            font.HighlightColor = Color.Yellow;

            doc.FirstSection.Body.FirstParagraph.AppendChild(run);
            doc.Save(ArtifactsDir + "Font.CreateFormattedRun.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.CreateFormattedRun.docx");
            run = doc.FirstSection.Body.FirstParagraph.Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Hello world!"));
            Assert.That(run.Font.Name, Is.EqualTo("Courier New"));
            Assert.That(run.Font.Size, Is.EqualTo(36));
            Assert.That(run.Font.HighlightColor.ToArgb(), Is.EqualTo(Color.Yellow.ToArgb()));
        }

        [Test]
        public void Caps()
        {
            //ExStart
            //ExFor:Font.AllCaps
            //ExFor:Font.SmallCaps
            //ExSummary:Shows how to format a run to display its contents in capitals.
            Document doc = new Document();
            Paragraph para = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            // There are two ways of getting a run to display its lowercase text in uppercase without changing the contents.
            // 1 -  Set the AllCaps flag to display all characters in regular capitals:
            Run run = new Run(doc, "all capitals");
            run.Font.AllCaps = true;
            para.AppendChild(run);

            para = (Paragraph)para.ParentNode.AppendChild(new Paragraph(doc));

            // 2 -  Set the SmallCaps flag to display all characters in small capitals:
            // If a character is lower case, it will appear in its upper case form
            // but will have the same height as the lower case (the font's x-height).
            // Characters that were in upper case originally will look the same.
            run = new Run(doc, "Small Capitals");
            run.Font.SmallCaps = true;
            para.AppendChild(run);

            doc.Save(ArtifactsDir + "Font.Caps.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Caps.docx");
            run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("all capitals"));
            Assert.That(run.Font.AllCaps, Is.True);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Small Capitals"));
            Assert.That(run.Font.SmallCaps, Is.True);
        }

        [Test]
        public void GetDocumentFonts()
        {
            //ExStart
            //ExFor:FontInfoCollection
            //ExFor:DocumentBase.FontInfos
            //ExFor:FontInfo
            //ExFor:FontInfo.Name
            //ExFor:FontInfo.IsTrueType
            //ExSummary:Shows how to print the details of what fonts are present in a document.
            Document doc = new Document(MyDir + "Embedded font.docx");

            FontInfoCollection allFonts = doc.FontInfos;
            Assert.That(allFonts.Count, Is.EqualTo(5)); //ExSkip

            // Print all the used and unused fonts in the document.
            for (int i = 0; i < allFonts.Count; i++)
            {
                Console.WriteLine($"Font index #{i}");
                Console.WriteLine($"\tName: {allFonts[i].Name}");
                Console.WriteLine($"\tIs {(allFonts[i].IsTrueType ? "" : "not ")}a trueType font");
            }
            //ExEnd
        }

        [Test]
        [Description("WORDSNET-16234")]
        public void DefaultValuesEmbeddedFontsParameters()
        {
            Document doc = new Document();

            Assert.That(doc.FontInfos.EmbedTrueTypeFonts, Is.False);
            Assert.That(doc.FontInfos.EmbedSystemFonts, Is.False);
            Assert.That(doc.FontInfos.SaveSubsetFonts, Is.False);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void FontInfoCollection(bool embedAllFonts)
        {
            //ExStart
            //ExFor:FontInfoCollection
            //ExFor:DocumentBase.FontInfos
            //ExFor:FontInfoCollection.EmbedTrueTypeFonts
            //ExFor:FontInfoCollection.EmbedSystemFonts
            //ExFor:FontInfoCollection.SaveSubsetFonts
            //ExSummary:Shows how to save a document with embedded TrueType fonts.
            Document doc = new Document(MyDir + "Document.docx");

            FontInfoCollection fontInfos = doc.FontInfos;
            fontInfos.EmbedTrueTypeFonts = embedAllFonts;
            fontInfos.EmbedSystemFonts = embedAllFonts;
            fontInfos.SaveSubsetFonts = embedAllFonts;

            doc.Save(ArtifactsDir + "Font.FontInfoCollection.docx");
            //ExEnd

            var testedFileLength = new FileInfo(ArtifactsDir + "Font.FontInfoCollection.docx").Length;

            if (embedAllFonts)
                Assert.That(testedFileLength < 28000, Is.True);
            else
                Assert.That(testedFileLength < 13000, Is.True);
        }

        [TestCase(true, false, false, Description =
            "Save a document with embedded TrueType fonts. System fonts are not included. Saves full versions of embedding fonts.")]
        [TestCase(true, true, false, Description =
            "Save a document with embedded TrueType fonts. System fonts are included. Saves full versions of embedding fonts.")]
        [TestCase(true, true, true, Description =
            "Save a document with embedded TrueType fonts. System fonts are included. Saves subset of embedding fonts.")]
        [TestCase(true, false, true, Description =
            "Save a document with embedded TrueType fonts. System fonts are not included. Saves subset of embedding fonts.")]
        [TestCase(false, false, false, Description = "Remove embedded fonts from the saved document.")]
        public void WorkWithEmbeddedFonts(bool embedTrueTypeFonts, bool embedSystemFonts, bool saveSubsetFonts)
        {
            Document doc = new Document(MyDir + "Document.docx");

            FontInfoCollection fontInfos = doc.FontInfos;
            fontInfos.EmbedTrueTypeFonts = embedTrueTypeFonts;
            fontInfos.EmbedSystemFonts = embedSystemFonts;
            fontInfos.SaveSubsetFonts = saveSubsetFonts;

            doc.Save(ArtifactsDir + "Font.WorkWithEmbeddedFonts.docx");
        }

        [Test]
        public void StrikeThrough()
        {
            //ExStart
            //ExFor:Font.StrikeThrough
            //ExFor:Font.DoubleStrikeThrough
            //ExSummary:Shows how to add a line strikethrough to text.
            Document doc = new Document();
            Paragraph para = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            Run run = new Run(doc, "Text with a single-line strikethrough.");
            run.Font.StrikeThrough = true;
            para.AppendChild(run);

            para = (Paragraph)para.ParentNode.AppendChild(new Paragraph(doc));

            run = new Run(doc, "Text with a double-line strikethrough.");
            run.Font.DoubleStrikeThrough = true;
            para.AppendChild(run);

            doc.Save(ArtifactsDir + "Font.StrikeThrough.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.StrikeThrough.docx");

            run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Text with a single-line strikethrough."));
            Assert.That(run.Font.StrikeThrough, Is.True);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Text with a double-line strikethrough."));
            Assert.That(run.Font.DoubleStrikeThrough, Is.True);
        }

        [Test]
        public void PositionSubscript()
        {
            //ExStart
            //ExFor:Font.Position
            //ExFor:Font.Subscript
            //ExFor:Font.Superscript
            //ExSummary:Shows how to format text to offset its position.
            Document doc = new Document();
            Paragraph para = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            // Raise this run of text 5 points above the baseline.
            Run run = new Run(doc, "Raised text. ");
            run.Font.Position = 5;
            para.AppendChild(run);

            // Lower this run of text 10 points below the baseline.
            run = new Run(doc, "Lowered text. ");
            run.Font.Position = -10;
            para.AppendChild(run);

            // Add a run of normal text.
            run = new Run(doc, "Text in its default position. ");
            para.AppendChild(run);

            // Add a run of text that appears as subscript.
            run = new Run(doc, "Subscript. ");
            run.Font.Subscript = true;
            para.AppendChild(run);

            // Add a run of text that appears as superscript.
            run = new Run(doc, "Superscript.");
            run.Font.Superscript = true;
            para.AppendChild(run);

            doc.Save(ArtifactsDir + "Font.PositionSubscript.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.PositionSubscript.docx");
            run = doc.FirstSection.Body.FirstParagraph.Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Raised text."));
            Assert.That(run.Font.Position, Is.EqualTo(5));

            doc = new Document(ArtifactsDir + "Font.PositionSubscript.docx");
            run = doc.FirstSection.Body.FirstParagraph.Runs[1];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Lowered text."));
            Assert.That(run.Font.Position, Is.EqualTo(-10));

            run = doc.FirstSection.Body.FirstParagraph.Runs[3];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Subscript."));
            Assert.That(run.Font.Subscript, Is.True);

            run = doc.FirstSection.Body.FirstParagraph.Runs[4];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Superscript."));
            Assert.That(run.Font.Superscript, Is.True);
        }

        [Test]
        public void ScalingSpacing()
        {
            //ExStart
            //ExFor:Font.Scaling
            //ExFor:Font.Spacing
            //ExSummary:Shows how to set horizontal scaling and spacing for characters.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add run of text and increase character width to 150%.
            builder.Font.Scaling = 150;
            builder.Writeln("Wide characters");

            // Add run of text and add 1pt of extra horizontal spacing between each character.
            builder.Font.Spacing = 1;
            builder.Writeln("Expanded by 1pt");

            // Add run of text and bring characters closer together by 1pt.
            builder.Font.Spacing = -1;
            builder.Writeln("Condensed by 1pt");

            doc.Save(ArtifactsDir + "Font.ScalingSpacing.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.ScalingSpacing.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Wide characters"));
            Assert.That(run.Font.Scaling, Is.EqualTo(150));

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Expanded by 1pt"));
            Assert.That(run.Font.Spacing, Is.EqualTo(1));

            run = doc.FirstSection.Body.Paragraphs[2].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Condensed by 1pt"));
            Assert.That(run.Font.Spacing, Is.EqualTo(-1));
        }

        [Test]
        public void Italic()
        {
            //ExStart
            //ExFor:Font.Italic
            //ExSummary:Shows how to write italicized text using a document builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Size = 36;
            builder.Font.Italic = true;
            builder.Writeln("Hello world!");

            doc.Save(ArtifactsDir + "Font.Italic.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Italic.docx");
            Run run = doc.FirstSection.Body.FirstParagraph.Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Hello world!"));
            Assert.That(run.Font.Italic, Is.True);
        }

        [Test]
        public void EngraveEmboss()
        {
            //ExStart
            //ExFor:Font.Emboss
            //ExFor:Font.Engrave
            //ExSummary:Shows how to apply engraving/embossing effects to text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Size = 36;
            builder.Font.Color = Color.LightBlue;

            // Below are two ways of using shadows to apply a 3D-like effect to the text.
            // 1 -  Engrave text to make it look like the letters are sunken into the page:
            builder.Font.Engrave = true;

            builder.Writeln("This text is engraved.");

            // 2 -  Emboss text to make it look like the letters pop out of the page:
            builder.Font.Engrave = false;
            builder.Font.Emboss = true;

            builder.Writeln("This text is embossed.");

            doc.Save(ArtifactsDir + "Font.EngraveEmboss.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.EngraveEmboss.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("This text is engraved."));
            Assert.That(run.Font.Engrave, Is.True);
            Assert.That(run.Font.Emboss, Is.False);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("This text is embossed."));
            Assert.That(run.Font.Engrave, Is.False);
            Assert.That(run.Font.Emboss, Is.True);
        }

        [Test]
        public void Shadow()
        {
            //ExStart
            //ExFor:Font.Shadow
            //ExSummary:Shows how to create a run of text formatted with a shadow.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the Shadow flag to apply an offset shadow effect,
            // making it look like the letters are floating above the page.
            builder.Font.Shadow = true;
            builder.Font.Size = 36;

            builder.Writeln("This text has a shadow.");

            doc.Save(ArtifactsDir + "Font.Shadow.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Shadow.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("This text has a shadow."));
            Assert.That(run.Font.Shadow, Is.True);
        }

        [Test]
        public void Outline()
        {
            //ExStart
            //ExFor:Font.Outline
            //ExSummary:Shows how to create a run of text formatted as outline.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the Outline flag to change the text's fill color to white and
            // leave a thin outline around each character in the original color of the text. 
            builder.Font.Outline = true;
            builder.Font.Color = Color.Blue;
            builder.Font.Size = 36;

            builder.Writeln("This text has an outline.");

            doc.Save(ArtifactsDir + "Font.Outline.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Outline.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("This text has an outline."));
            Assert.That(run.Font.Outline, Is.True);
        }

        [Test]
        public void Hidden()
        {
            //ExStart
            //ExFor:Font.Hidden
            //ExSummary:Shows how to create a run of hidden text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // With the Hidden flag set to true, any text that we create using this Font object will be invisible in the document.
            // We will not see or highlight hidden text unless we enable the "Hidden text" option
            // found in Microsoft Word via "File" -> "Options" -> "Display". The text will still be there,
            // and we will be able to access this text programmatically.
            // It is not advised to use this method to hide sensitive information.
            builder.Font.Hidden = true;
            builder.Font.Size = 36;

            builder.Writeln("This text will not be visible in the document.");

            doc.Save(ArtifactsDir + "Font.Hidden.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Hidden.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("This text will not be visible in the document."));
            Assert.That(run.Font.Hidden, Is.True);
        }

        [Test]
        public void Kerning()
        {
            //ExStart
            //ExFor:Font.Kerning
            //ExSummary:Shows how to specify the font size at which kerning begins to take effect.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Arial Black";

            // Set the builder's font size, and minimum size at which kerning will take effect.
            // The font size falls below the kerning threshold, so the run bellow will not have kerning.
            builder.Font.Size = 18;
            builder.Font.Kerning = 24;

            builder.Writeln("TALLY. (Kerning not applied)");

            // Set the kerning threshold so that the builder's current font size is above it.
            // Any text we add from this point will have kerning applied. The spaces between characters
            // will be adjusted, normally resulting in a slightly more aesthetically pleasing text run.
            builder.Font.Kerning = 12;

            builder.Writeln("TALLY. (Kerning applied)");

            doc.Save(ArtifactsDir + "Font.Kerning.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Kerning.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("TALLY. (Kerning not applied)"));
            Assert.That(run.Font.Kerning, Is.EqualTo(24));
            Assert.That(run.Font.Size, Is.EqualTo(18));

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("TALLY. (Kerning applied)"));
            Assert.That(run.Font.Kerning, Is.EqualTo(12));
            Assert.That(run.Font.Size, Is.EqualTo(18));
        }

        [Test]
        public void NoProofing()
        {
            //ExStart
            //ExFor:Font.NoProofing
            //ExSummary:Shows how to prevent text from being spell checked by Microsoft Word.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Normally, Microsoft Word emphasizes spelling errors with a jagged red underline.
            // We can un-set the "NoProofing" flag to create a portion of text that
            // bypasses the spell checker while completely disabling it.
            builder.Font.NoProofing = true;

            builder.Writeln("Proofing has been disabled, so these spelking errrs will not display red lines underneath.");

            doc.Save(ArtifactsDir + "Font.NoProofing.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.NoProofing.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Proofing has been disabled, so these spelking errrs will not display red lines underneath."));
            Assert.That(run.Font.NoProofing, Is.True);
        }

        [Test]
        public void LocaleId()
        {
            //ExStart
            //ExFor:Font.LocaleId
            //ExSummary:Shows how to set the locale of the text that we are adding with a document builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // If we set the font's locale to English and insert some Russian text,
            // the English locale spell checker will not recognize the text and detect it as a spelling error.
            builder.Font.LocaleId = new CultureInfo("en-US", false).LCID;
            builder.Writeln("Привет!");

            // Set a matching locale for the text that we are about to add to apply the appropriate spell checker.
            builder.Font.LocaleId = new CultureInfo("ru-RU", false).LCID;
            builder.Writeln("Привет!");

            doc.Save(ArtifactsDir + "Font.LocaleId.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.LocaleId.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Привет!"));
            Assert.That(run.Font.LocaleId, Is.EqualTo(1033));

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Привет!"));
            Assert.That(run.Font.LocaleId, Is.EqualTo(1049));
        }

        [Test]
        public void Underlines()
        {
            //ExStart
            //ExFor:Font.Underline
            //ExFor:Font.UnderlineColor
            //ExSummary:Shows how to configure the style and color of a text underline.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Underline = Underline.Dotted;
            builder.Font.UnderlineColor = Color.Red;

            builder.Writeln("Underlined text.");

            doc.Save(ArtifactsDir + "Font.Underlines.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Underlines.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Underlined text."));
            Assert.That(run.Font.Underline, Is.EqualTo(Underline.Dotted));
            Assert.That(run.Font.UnderlineColor.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
        }

        [Test]
        public void ComplexScript()
        {
            //ExStart
            //ExFor:Font.ComplexScript
            //ExSummary:Shows how to add text that is always treated as complex script.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.ComplexScript = true;

            builder.Writeln("Text treated as complex script.");

            doc.Save(ArtifactsDir + "Font.ComplexScript.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.ComplexScript.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Text treated as complex script."));
            Assert.That(run.Font.ComplexScript, Is.True);
        }

        [Test]
        public void SparklingText()
        {
            //ExStart
            //ExFor:TextEffect
            //ExFor:Font.TextEffect
            //ExSummary:Shows how to apply a visual effect to a run.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Size = 36;
            builder.Font.TextEffect = TextEffect.SparkleText;

            builder.Writeln("Text with a sparkle effect.");

            // Older versions of Microsoft Word only support font animation effects.
            doc.Save(ArtifactsDir + "Font.SparklingText.doc");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.SparklingText.doc");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Text with a sparkle effect."));
            Assert.That(run.Font.TextEffect, Is.EqualTo(TextEffect.SparkleText));
        }

        [Test]
        public void ForegroundAndBackground()
        {
            //ExStart
            //ExFor:Shading.ForegroundPatternThemeColor
            //ExFor:Shading.BackgroundPatternThemeColor
            //ExFor:Shading.ForegroundTintAndShade
            //ExFor:Shading.BackgroundTintAndShade
            //ExSummary:Shows how to set foreground and background colors for shading texture.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shading shading = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Shading;
            shading.Texture = TextureIndex.Texture12Pt5Percent;
            shading.ForegroundPatternThemeColor = ThemeColor.Dark1;
            shading.BackgroundPatternThemeColor = ThemeColor.Dark2;

            shading.ForegroundTintAndShade = 0.5;
            shading.BackgroundTintAndShade = -0.2;

            builder.Font.Border.Color = Color.Green;
            builder.Font.Border.LineWidth = 2.5d;
            builder.Font.Border.LineStyle = LineStyle.DashDotStroker;

            builder.Writeln("Foreground and background pattern colors for shading texture.");

            doc.Save(ArtifactsDir + "Font.ForegroundAndBackground.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.ForegroundAndBackground.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Foreground and background pattern colors for shading texture."));
            Assert.That(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.Shading.ForegroundPatternThemeColor, Is.EqualTo(ThemeColor.Dark1));
            Assert.That(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.Shading.BackgroundPatternThemeColor, Is.EqualTo(ThemeColor.Dark2));

            Assert.That(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.Shading.ForegroundTintAndShade, Is.EqualTo(0.5).Within(0.1));
            Assert.That(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.Shading.BackgroundTintAndShade, Is.EqualTo(-0.2).Within(0.1));
        }

        [Test]
        public void Shading()
        {
            //ExStart
            //ExFor:Font.Shading
            //ExSummary:Shows how to apply shading to text created by a document builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Color = Color.White;

            // One way to make the text created using our white font color visible
            // is to apply a background shading effect.
            Shading shading = builder.Font.Shading;
            shading.Texture = TextureIndex.TextureDiagonalUp;
            shading.BackgroundPatternColor = Color.OrangeRed;
            shading.ForegroundPatternColor = Color.DarkBlue;

            builder.Writeln("White text on an orange background with a two-tone texture.");

            doc.Save(ArtifactsDir + "Font.Shading.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Shading.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("White text on an orange background with a two-tone texture."));
            Assert.That(run.Font.Color.ToArgb(), Is.EqualTo(Color.White.ToArgb()));

            Assert.That(run.Font.Shading.Texture, Is.EqualTo(TextureIndex.TextureDiagonalUp));
            Assert.That(run.Font.Shading.BackgroundPatternColor.ToArgb(), Is.EqualTo(Color.OrangeRed.ToArgb()));
            Assert.That(run.Font.Shading.ForegroundPatternColor.ToArgb(), Is.EqualTo(Color.DarkBlue.ToArgb()));
        }

        [Test, Category("SkipMono")]
        public void Bidi()
        {
            //ExStart
            //ExFor:Font.Bidi
            //ExFor:Font.NameBi
            //ExFor:Font.SizeBi
            //ExFor:Font.ItalicBi
            //ExFor:Font.BoldBi
            //ExFor:Font.LocaleIdBi
            //ExSummary:Shows how to define separate sets of font settings for right-to-left, and right-to-left text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define a set of font settings for left-to-right text.
            builder.Font.Name = "Courier New";
            builder.Font.Size = 16;
            builder.Font.Italic = false;
            builder.Font.Bold = false;
            builder.Font.LocaleId = new CultureInfo("en-US", false).LCID;

            // Define another set of font settings for right-to-left text.
            builder.Font.NameBi = "Andalus";
            builder.Font.SizeBi = 24;
            builder.Font.ItalicBi = true;
            builder.Font.BoldBi = true;
            builder.Font.LocaleIdBi = new CultureInfo("ar-AR", false).LCID;

            // We can use the Bidi flag to indicate whether the text we are about to add
            // with the document builder is right-to-left. When we add text with this flag set to true,
            // it will be formatted using the right-to-left set of font settings.
            builder.Font.Bidi = true;
            builder.Write("مرحبًا");

            // Set the flag to false, and then add left-to-right text.
            // The document builder will format these using the left-to-right set of font settings.
            builder.Font.Bidi = false;
            builder.Write(" Hello world!");

            doc.Save(ArtifactsDir + "Font.Bidi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Bidi.docx");

            foreach (Run run in doc.FirstSection.Body.Paragraphs[0].Runs)
            {
                switch (doc.FirstSection.Body.Paragraphs[0].IndexOf(run))
                {
                    case 0:
                        Assert.That(run.GetText().Trim(), Is.EqualTo("مرحبًا"));
                        Assert.That(run.Font.Bidi, Is.True);
                        break;
                    case 1:
                        Assert.That(run.GetText().Trim(), Is.EqualTo("Hello world!"));
                        Assert.That(run.Font.Bidi, Is.False);
                        break;
                }

                Assert.That(run.Font.LocaleId, Is.EqualTo(1033));
                Assert.That(run.Font.Size, Is.EqualTo(16));
                Assert.That(run.Font.Italic, Is.False);
                Assert.That(run.Font.Bold, Is.False);
                Assert.That(run.Font.LocaleIdBi, Is.EqualTo(1025));
                Assert.That(run.Font.SizeBi, Is.EqualTo(24));
                Assert.That(run.Font.NameBi, Is.EqualTo("Andalus"));
                Assert.That(run.Font.ItalicBi, Is.True);
                Assert.That(run.Font.BoldBi, Is.True);
            }
        }

        [Test]
        public void FarEast()
        {
            //ExStart
            //ExFor:Font.NameFarEast
            //ExFor:Font.LocaleIdFarEast
            //ExSummary:Shows how to insert and format text in a Far East language.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font settings that the document builder will apply to any text that it inserts.
            builder.Font.Name = "Courier New";
            builder.Font.LocaleId = new CultureInfo("en-US", false).LCID;

            // Name "FarEast" equivalents for our font and locale.
            // If the builder inserts Asian characters with this Font configuration, then each run that contains
            // these characters will display them using the "FarEast" font/locale instead of the default.
            // This could be useful when a western font does not have ideal representations for Asian characters.
            builder.Font.NameFarEast = "SimSun";
            builder.Font.LocaleIdFarEast = new CultureInfo("zh-CN", false).LCID;

            // This text will be displayed in the default font/locale.
            builder.Writeln("Hello world!");

            // Since these are Asian characters, this run will apply our "FarEast" font/locale equivalents.
            builder.Writeln("你好世界");

            doc.Save(ArtifactsDir + "Font.FarEast.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.FarEast.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Hello world!"));
            Assert.That(run.Font.LocaleId, Is.EqualTo(1033));
            Assert.That(run.Font.Name, Is.EqualTo("Courier New"));
            Assert.That(run.Font.LocaleIdFarEast, Is.EqualTo(2052));
            Assert.That(run.Font.NameFarEast, Is.EqualTo("SimSun"));

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("你好世界"));
            Assert.That(run.Font.LocaleId, Is.EqualTo(1033));
            Assert.That(run.Font.Name, Is.EqualTo("SimSun"));
            Assert.That(run.Font.LocaleIdFarEast, Is.EqualTo(2052));
            Assert.That(run.Font.NameFarEast, Is.EqualTo("SimSun"));
        }

        [Test]
        public void NameAscii()
        {
            //ExStart
            //ExFor:Font.NameAscii
            //ExFor:Font.NameOther
            //ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Suppose a run that we use the builder to insert while using this font configuration
            // contains characters within the ASCII characters' range. In that case,
            // it will display those characters using this font.
            builder.Font.NameAscii = "Calibri";

            // With no other font specified, the builder will also apply this font to all characters that it inserts.
            Assert.That(builder.Font.Name, Is.EqualTo("Calibri"));

            // Specify a font to use for all characters outside of the ASCII range.
            // Ideally, this font should have a glyph for each required non-ASCII character code.
            builder.Font.NameOther = "Courier New";

            // Insert a run with one word consisting of ASCII characters, and one word with all characters outside that range.
            // Each character will be displayed using either of the fonts, depending on.
            builder.Writeln("Hello, Привет");

            doc.Save(ArtifactsDir + "Font.NameAscii.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.NameAscii.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("Hello, Привет"));
            Assert.That(run.Font.Name, Is.EqualTo("Calibri"));
            Assert.That(run.Font.NameAscii, Is.EqualTo("Calibri"));
            Assert.That(run.Font.NameOther, Is.EqualTo("Courier New"));
        }

        [Test]
        public void ChangeStyle()
        {
            //ExStart
            //ExFor:Font.StyleName
            //ExFor:Font.StyleIdentifier
            //ExFor:StyleIdentifier
            //ExSummary:Shows how to change the style of existing text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two ways of referencing styles.
            // 1 -  Using the style name:
            builder.Font.StyleName = "Emphasis";
            builder.Writeln("Text originally in \"Emphasis\" style");

            // 2 -  Using a built-in style identifier:
            builder.Font.StyleIdentifier = StyleIdentifier.IntenseEmphasis;
            builder.Writeln("Text originally in \"Intense Emphasis\" style");

            // Convert all uses of one style to another,
            // using the above methods to reference old and new styles.
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                if (run.Font.StyleName == "Emphasis")
                    run.Font.StyleName = "Strong";

                if (run.Font.StyleIdentifier == StyleIdentifier.IntenseEmphasis)
                    run.Font.StyleIdentifier = StyleIdentifier.Strong;
            }

            doc.Save(ArtifactsDir + "Font.ChangeStyle.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.ChangeStyle.docx");
            Run docRun = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(docRun.GetText().Trim(), Is.EqualTo("Text originally in \"Emphasis\" style"));
            Assert.That(docRun.Font.StyleIdentifier, Is.EqualTo(StyleIdentifier.Strong));
            Assert.That(docRun.Font.StyleName, Is.EqualTo("Strong"));

            docRun = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(docRun.GetText().Trim(), Is.EqualTo("Text originally in \"Intense Emphasis\" style"));
            Assert.That(docRun.Font.StyleIdentifier, Is.EqualTo(StyleIdentifier.Strong));
            Assert.That(docRun.Font.StyleName, Is.EqualTo("Strong"));
        }

        [Test]
        public void BuiltIn()
        {
            //ExStart
            //ExFor:Style.BuiltIn
            //ExSummary:Shows how to differentiate custom styles from built-in styles.
            Document doc = new Document();

            // When we create a document using Microsoft Word, or programmatically using Aspose.Words,
            // the document will come with a collection of styles to apply to its text to modify its appearance.
            // We can access these built-in styles via the document's "Styles" collection.
            // These styles will all have the "BuiltIn" flag set to "true".
            Style style = doc.Styles["Emphasis"];

            Assert.That(style.BuiltIn, Is.True);

            // Create a custom style and add it to the collection.
            // Custom styles such as this will have the "BuiltIn" flag set to "false". 
            style = doc.Styles.Add(StyleType.Character, "MyStyle");
            style.Font.Color = Color.Navy;
            style.Font.Name = "Courier New";

            Assert.That(style.BuiltIn, Is.False);
            //ExEnd
        }

        [Test]
        public void Style()
        {
            //ExStart
            //ExFor:Font.Style
            //ExSummary:Applies a double underline to all runs in a document that are formatted with custom character styles.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a custom style and apply it to text created using a document builder.
            Style style = doc.Styles.Add(StyleType.Character, "MyStyle");
            style.Font.Color = Color.Red;
            style.Font.Name = "Courier New";

            builder.Font.StyleName = "MyStyle";
            builder.Write("This text is in a custom style.");

            // Iterate over every run and add a double underline to every custom style.
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                Style charStyle = run.Font.Style;

                if (!charStyle.BuiltIn)
                    run.Font.Underline = Underline.Double;
            }

            doc.Save(ArtifactsDir + "Font.Style.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Style.docx");
            Run docRun = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(docRun.GetText().Trim(), Is.EqualTo("This text is in a custom style."));
            Assert.That(docRun.Font.StyleName, Is.EqualTo("MyStyle"));
            Assert.That(docRun.Font.Style.BuiltIn, Is.False);
            Assert.That(docRun.Font.Underline, Is.EqualTo(Underline.Double));
        }

        [Test]
        public void GetAvailableFonts()
        {
            //ExStart
            //ExFor:PhysicalFontInfo
            //ExFor:FontSourceBase.GetAvailableFonts
            //ExFor:PhysicalFontInfo.FontFamilyName
            //ExFor:PhysicalFontInfo.FullFontName
            //ExFor:PhysicalFontInfo.Version
            //ExFor:PhysicalFontInfo.FilePath
            //ExSummary:Shows how to list available fonts.
            // Configure Aspose.Words to source fonts from a custom folder, and then print every available font.
            FontSourceBase[] folderFontSource = { new FolderFontSource(FontsDir, true) };

            foreach (PhysicalFontInfo fontInfo in folderFontSource[0].GetAvailableFonts())
            {
                Console.WriteLine("FontFamilyName : {0}", fontInfo.FontFamilyName);
                Console.WriteLine("FullFontName  : {0}", fontInfo.FullFontName);
                Console.WriteLine("Version  : {0}", fontInfo.Version);
                Console.WriteLine("FilePath : {0}\n", fontInfo.FilePath);
            }
            //ExEnd

            Assert.That(Directory.EnumerateFiles(FontsDir, "*.*", SearchOption.AllDirectories).Count(f => f.EndsWith(".ttf") || f.EndsWith(".otf")), Is.EqualTo(folderFontSource[0].GetAvailableFonts().Count));
        }

        [Test]
        public void SetFontAutoColor()
        {
            //ExStart
            //ExFor:Font.AutoColor
            //ExSummary:Shows how to improve readability by automatically selecting text color based on the brightness of its background.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // If a run's Font object does not specify text color, it will automatically
            // select either black or white depending on the background color's color.
            Assert.That(builder.Font.Color.ToArgb(), Is.EqualTo(Color.Empty.ToArgb()));

            // The default color for text is black. If the color of the background is dark, black text will be difficult to see.
            // To solve this problem, the AutoColor property will display this text in white.
            builder.Font.Shading.BackgroundPatternColor = Color.DarkBlue;

            builder.Writeln("The text color automatically chosen for this run is white.");

            Assert.That(doc.FirstSection.Body.Paragraphs[0].Runs[0].Font.AutoColor.ToArgb(), Is.EqualTo(Color.White.ToArgb()));

            // If we change the background to a light color, black will be a more
            // suitable text color than white so that the auto color will display it in black.
            builder.Font.Shading.BackgroundPatternColor = Color.LightBlue;

            builder.Writeln("The text color automatically chosen for this run is black.");

            Assert.That(doc.FirstSection.Body.Paragraphs[1].Runs[0].Font.AutoColor.ToArgb(), Is.EqualTo(Color.Black.ToArgb()));

            doc.Save(ArtifactsDir + "Font.SetFontAutoColor.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.SetFontAutoColor.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("The text color automatically chosen for this run is white."));
            Assert.That(run.Font.Color.ToArgb(), Is.EqualTo(Color.Empty.ToArgb()));
            Assert.That(run.Font.Shading.BackgroundPatternColor.ToArgb(), Is.EqualTo(Color.DarkBlue.ToArgb()));

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.That(run.GetText().Trim(), Is.EqualTo("The text color automatically chosen for this run is black."));
            Assert.That(run.Font.Color.ToArgb(), Is.EqualTo(Color.Empty.ToArgb()));
            Assert.That(run.Font.Shading.BackgroundPatternColor.ToArgb(), Is.EqualTo(Color.LightBlue.ToArgb()));
        }

        //ExStart
        //ExFor:Font.Hidden
        //ExFor:Paragraph.Accept(DocumentVisitor)
        //ExFor:Paragraph.AcceptStart(DocumentVisitor)
        //ExFor:Paragraph.AcceptEnd(DocumentVisitor)
        //ExFor:DocumentVisitor.VisitParagraphStart(Paragraph)
        //ExFor:DocumentVisitor.VisitFormField(FormField)
        //ExFor:DocumentVisitor.VisitTableEnd(Table)
        //ExFor:DocumentVisitor.VisitCellEnd(Cell)
        //ExFor:DocumentVisitor.VisitRowEnd(Row)
        //ExFor:DocumentVisitor.VisitSpecialChar(SpecialChar)
        //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
        //ExFor:DocumentVisitor.VisitShapeStart(Shape)
        //ExFor:DocumentVisitor.VisitCommentStart(Comment)
        //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
        //ExFor:SpecialChar
        //ExFor:SpecialChar.Accept(DocumentVisitor)
        //ExFor:SpecialChar.GetText
        //ExFor:Node.Accept(DocumentVisitor)
        //ExFor:Paragraph.ParagraphBreakFont
        //ExFor:Table.Accept(DocumentVisitor)
        //ExFor:Table.AcceptStart(DocumentVisitor)
        //ExFor:Table.AcceptEnd(DocumentVisitor)
        //ExSummary:Shows how to use a DocumentVisitor implementation to remove all hidden content from a document.
        [Test] //ExSkip
        public void RemoveHiddenContentFromDocument()
        {
            Document doc = new Document(MyDir + "Hidden content.docx");
            Assert.That(doc.GetChildNodes(NodeType.Paragraph, true).Count, Is.EqualTo(26)); //ExSkip
            Assert.That(doc.GetChildNodes(NodeType.Table, true).Count, Is.EqualTo(2)); //ExSkip

            RemoveHiddenContentVisitor hiddenContentRemover = new RemoveHiddenContentVisitor();

            // Below are three types of fields which can accept a document visitor,
            // which will allow it to visit the accepting node, and then traverse its child nodes in a depth-first manner.
            // 1 -  Paragraph node:
            Paragraph para = (Paragraph)doc.GetChild(NodeType.Paragraph, 4, true);
            para.Accept(hiddenContentRemover);

            // 2 -  Table node:
            Table table = doc.FirstSection.Body.Tables[0];
            table.Accept(hiddenContentRemover);

            // 3 -  Document node:
            doc.Accept(hiddenContentRemover);

            doc.Save(ArtifactsDir + "Font.RemoveHiddenContentFromDocument.docx");
            TestRemoveHiddenContent(new Document(ArtifactsDir + "Font.RemoveHiddenContentFromDocument.docx")); //ExSkip
        }

        /// <summary>
        /// Removes all visited nodes marked as "hidden content".
        /// </summary>
        public class RemoveHiddenContentVisitor : DocumentVisitor
        {
            /// <summary>
            /// Called when a FieldStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldStart(FieldStart fieldStart)
            {
                if (fieldStart.Font.Hidden)
                    fieldStart.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                if (fieldEnd.Font.Hidden)
                    fieldEnd.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                if (fieldSeparator.Font.Hidden)
                    fieldSeparator.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (run.Font.Hidden)
                    run.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Paragraph node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitParagraphStart(Paragraph paragraph)
            {
                if (paragraph.ParagraphBreakFont.Hidden)
                    paragraph.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FormField is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFormField(FormField formField)
            {
                if (formField.Font.Hidden)
                    formField.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a GroupShape is encountered in the document.
            /// </summary>
            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                if (groupShape.Font.Hidden)
                    groupShape.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Shape is encountered in the document.
            /// </summary>
            public override VisitorAction VisitShapeStart(Shape shape)
            {
                if (shape.Font.Hidden)
                    shape.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Comment is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentStart(Comment comment)
            {
                if (comment.Font.Hidden)
                    comment.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Footnote is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFootnoteStart(Footnote footnote)
            {
                if (footnote.Font.Hidden)
                    footnote.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a SpecialCharacter is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSpecialChar(SpecialChar specialChar)
            {
                Console.WriteLine(specialChar.GetText());

                if (specialChar.Font.Hidden)
                    specialChar.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Table node is ended in the document.
            /// </summary>
            public override VisitorAction VisitTableEnd(Table table)
            {
                // The content inside table cells may have the hidden content flag, but the tables themselves cannot.
                // If this table had nothing but hidden content, this visitor would have removed all of it,
                // and there would be no child nodes left.
                // Thus, we can also treat the table itself as hidden content and remove it.
                // Tables which are empty but do not have hidden content will have cells with empty paragraphs inside,
                // which this visitor will not remove.
                if (!table.HasChildNodes)
                    table.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Cell node is ended in the document.
            /// </summary>
            public override VisitorAction VisitCellEnd(Cell cell)
            {
                if (!cell.HasChildNodes && cell.ParentNode != null)
                    cell.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Row node is ended in the document.
            /// </summary>
            public override VisitorAction VisitRowEnd(Row row)
            {
                if (!row.HasChildNodes && row.ParentNode != null)
                    row.Remove();

                return VisitorAction.Continue;
            }
        }
        //ExEnd

        private void TestRemoveHiddenContent(Document doc)
        {
            Assert.That(doc.GetChildNodes(NodeType.Paragraph, true).Count, Is.EqualTo(20)); //ExSkip
            Assert.That(doc.GetChildNodes(NodeType.Table, true).Count, Is.EqualTo(1)); //ExSkip

            foreach (Node node in doc.GetChildNodes(NodeType.Any, true))
            {
                switch (node)
                {
                    case FieldStart fieldStart:
                        Assert.That(fieldStart.Font.Hidden, Is.False);
                        break;
                    case FieldEnd fieldEnd:
                        Assert.That(fieldEnd.Font.Hidden, Is.False);
                        break;
                    case FieldSeparator fieldSeparator:
                        Assert.That(fieldSeparator.Font.Hidden, Is.False);
                        break;
                    case Run run:
                        Assert.That(run.Font.Hidden, Is.False);
                        break;
                    case Paragraph paragraph:
                        Assert.That(paragraph.ParagraphBreakFont.Hidden, Is.False);
                        break;
                    case FormField formField:
                        Assert.That(formField.Font.Hidden, Is.False);
                        break;
                    case GroupShape groupShape:
                        Assert.That(groupShape.Font.Hidden, Is.False);
                        break;
                    case Shape shape:
                        Assert.That(shape.Font.Hidden, Is.False);
                        break;
                    case Comment comment:
                        Assert.That(comment.Font.Hidden, Is.False);
                        break;
                    case Footnote footnote:
                        Assert.That(footnote.Font.Hidden, Is.False);
                        break;
                    case SpecialChar specialChar:
                        Assert.That(specialChar.Font.Hidden, Is.False);
                        break;
                }
            }
        }

        [Test]
        public void DefaultFonts()
        {
            //ExStart
            //ExFor:FontInfoCollection.Contains(String)
            //ExFor:FontInfoCollection.Count
            //ExSummary:Shows info about the fonts that are present in the blank document.
            Document doc = new Document();

            // A blank document contains 3 default fonts. Each font in the document
            // will have a corresponding FontInfo object which contains details about that font.
            Assert.That(doc.FontInfos.Count, Is.EqualTo(3));

            Assert.That(doc.FontInfos.Contains("Times New Roman"), Is.True);
            Assert.That(doc.FontInfos["Times New Roman"].Charset, Is.EqualTo(204));

            Assert.That(doc.FontInfos.Contains("Symbol"), Is.True);
            Assert.That(doc.FontInfos.Contains("Arial"), Is.True);
            //ExEnd
        }

        [Test]
        public void ExtractEmbeddedFont()
        {
            //ExStart
            //ExFor:EmbeddedFontFormat
            //ExFor:EmbeddedFontStyle
            //ExFor:FontInfo.GetEmbeddedFont(EmbeddedFontFormat,EmbeddedFontStyle)
            //ExFor:FontInfo.GetEmbeddedFontAsOpenType(EmbeddedFontStyle)
            //ExFor:FontInfoCollection.Item(Int32)
            //ExFor:FontInfoCollection.Item(String)
            //ExSummary:Shows how to extract an embedded font from a document, and save it to the local file system.
            Document doc = new Document(MyDir + "Embedded font.docx");

            FontInfo embeddedFont = doc.FontInfos["Alte DIN 1451 Mittelschrift"];
            byte[] embeddedFontBytes = embeddedFont.GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular);
            Assert.That(embeddedFontBytes, Is.Not.Null); //ExSkip

            File.WriteAllBytes(ArtifactsDir + "Alte DIN 1451 Mittelschrift.ttf", embeddedFontBytes);

            // Embedded font formats may be different in other formats such as .doc.
            // We need to know the correct format before we can extract the font.
            doc = new Document(MyDir + "Embedded font.doc");

            Assert.That(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular), Is.Null);
            Assert.That(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFont(EmbeddedFontFormat.EmbeddedOpenType, EmbeddedFontStyle.Regular), Is.Not.Null);

            // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType.
            embeddedFontBytes = doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFontAsOpenType(EmbeddedFontStyle.Regular);

            File.WriteAllBytes(ArtifactsDir + "Alte DIN 1451 Mittelschrift.otf", embeddedFontBytes);
            //ExEnd
        }

        [Test]
        public void GetFontInfoFromFile()
        {
            //ExStart
            //ExFor:FontFamily
            //ExFor:FontPitch
            //ExFor:FontInfo.AltName
            //ExFor:FontInfo.Charset
            //ExFor:FontInfo.Family
            //ExFor:FontInfo.Panose
            //ExFor:FontInfo.Pitch
            //ExFor:FontInfoCollection.GetEnumerator
            //ExSummary:Shows how to access and print details of each font in a document.
            Document doc = new Document(MyDir + "Document.docx");

            IEnumerator<FontInfo> fontCollectionEnumerator = doc.FontInfos.GetEnumerator();
            while (fontCollectionEnumerator.MoveNext())
            {
                FontInfo fontInfo = fontCollectionEnumerator.Current;
                if (fontInfo != null)
                {
                    Console.WriteLine("Font name: " + fontInfo.Name);

                    // Alt names are usually blank.
                    Console.WriteLine("Alt name: " + fontInfo.AltName);
                    Console.WriteLine("\t- Family: " + fontInfo.Family);
                    Console.WriteLine("\t- " + (fontInfo.IsTrueType ? "Is TrueType" : "Is not TrueType"));
                    Console.WriteLine("\t- Pitch: " + fontInfo.Pitch);
                    Console.WriteLine("\t- Charset: " + fontInfo.Charset);
                    Console.WriteLine("\t- Panose:");
                    Console.WriteLine("\t\tFamily Kind: " + fontInfo.Panose[0]);
                    Console.WriteLine("\t\tSerif Style: " + fontInfo.Panose[1]);
                    Console.WriteLine("\t\tWeight: " + fontInfo.Panose[2]);
                    Console.WriteLine("\t\tProportion: " + fontInfo.Panose[3]);
                    Console.WriteLine("\t\tContrast: " + fontInfo.Panose[4]);
                    Console.WriteLine("\t\tStroke Variation: " + fontInfo.Panose[5]);
                    Console.WriteLine("\t\tArm Style: " + fontInfo.Panose[6]);
                    Console.WriteLine("\t\tLetterform: " + fontInfo.Panose[7]);
                    Console.WriteLine("\t\tMidline: " + fontInfo.Panose[8]);
                    Console.WriteLine("\t\tX-Height: " + fontInfo.Panose[9]);
                }
            }
            //ExEnd

            Assert.That(doc.FontInfos["Calibri"].Panose, Is.EqualTo(new[] { 2, 15, 5, 2, 2, 2, 4, 3, 2, 4 }));
            Assert.That(doc.FontInfos["Calibri Light"].Panose, Is.EqualTo(new[] { 2, 15, 3, 2, 2, 2, 4, 3, 2, 4 }));
            Assert.That(doc.FontInfos["Times New Roman"].Panose, Is.EqualTo(new[] { 2, 2, 6, 3, 5, 4, 5, 2, 3, 4 }));
        }

        [Test]
        public void LineSpacing()
        {
            //ExStart
            //ExFor:Font.LineSpacing
            //ExSummary:Shows how to get a font's line spacing, in points.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set different fonts for the DocumentBuilder and verify their line spacing.
            builder.Font.Name = "Calibri";
            Assert.That(builder.Font.LineSpacing, Is.EqualTo(14.6484375d));

            builder.Font.Name = "Times New Roman";
            Assert.That(builder.Font.LineSpacing, Is.EqualTo(13.798828125d));
            //ExEnd
        }

        [Test]
        public void HasDmlEffect()
        {
            //ExStart
            //ExFor:Font.HasDmlEffect(TextDmlEffect)
            //ExFor:TextDmlEffect
            //ExSummary:Shows how to check if a run displays a DrawingML text effect.
            Document doc = new Document(MyDir + "DrawingML text effects.docx");

            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;

            Assert.That(runs[0].Font.HasDmlEffect(TextDmlEffect.Shadow), Is.True);
            Assert.That(runs[1].Font.HasDmlEffect(TextDmlEffect.Shadow), Is.True);
            Assert.That(runs[2].Font.HasDmlEffect(TextDmlEffect.Reflection), Is.True);
            Assert.That(runs[3].Font.HasDmlEffect(TextDmlEffect.Effect3D), Is.True);
            Assert.That(runs[4].Font.HasDmlEffect(TextDmlEffect.Fill), Is.True);
            //ExEnd
        }

        [Test, Category("SkipGitHub")]
        public void CheckScanUserFontsFolder()
        {
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var currentUserFontsFolder = Path.Combine(userProfile, @"AppData\Local\Microsoft\Windows\Fonts");
            var currentUserFonts = Directory.GetFiles(currentUserFontsFolder, "*.ttf");
            if (currentUserFonts.Length != 0)
            {
                // On Windows 10 fonts may be installed either into system folder "%windir%\fonts" for all users
                // or into user folder "%userprofile%\AppData\Local\Microsoft\Windows\Fonts" for current user.
                SystemFontSource systemFontSource = new SystemFontSource();
                Assert.That(systemFontSource.GetAvailableFonts()
                        .FirstOrDefault(x => x.FilePath.Contains("\\AppData\\Local\\Microsoft\\Windows\\Fonts")), Is.Not.Null, "Fonts did not install to the user font folder");
            }
        }

        [TestCase(EmphasisMark.None)]
        [TestCase(EmphasisMark.OverComma)]
        [TestCase(EmphasisMark.OverSolidCircle)]
        [TestCase(EmphasisMark.OverWhiteCircle)]
        [TestCase(EmphasisMark.UnderSolidCircle)]
        public void SetEmphasisMark(EmphasisMark emphasisMark)
        {
            //ExStart
            //ExFor:EmphasisMark
            //ExFor:Font.EmphasisMark
            //ExSummary:Shows how to add additional character rendered above/below the glyph-character.
            DocumentBuilder builder = new DocumentBuilder();

            // Possible types of emphasis mark:
            // https://apireference.aspose.com/words/net/aspose.words/emphasismark
            builder.Font.EmphasisMark = emphasisMark; 
            
            builder.Write("Emphasis text");
            builder.Writeln();
            builder.Font.ClearFormatting();
            builder.Write("Simple text");
 
            builder.Document.Save(ArtifactsDir + "Fonts.SetEmphasisMark.docx");
            //ExEnd
        }

        [Test]
        public void ThemeFontsColors()
        {
            //ExStart
            //ExFor:Font.ThemeFont
            //ExFor:Font.ThemeFontAscii
            //ExFor:Font.ThemeFontBi
            //ExFor:Font.ThemeFontFarEast
            //ExFor:Font.ThemeFontOther
            //ExFor:Font.ThemeColor
            //ExFor:ThemeFont
            //ExFor:ThemeColor
            //ExSummary:Shows how to work with theme fonts and colors.
            Document doc = new Document();

            // Define fonts for languages uses by default.
            doc.Theme.MinorFonts.Latin = "Algerian";
            doc.Theme.MinorFonts.EastAsian = "Aharoni";
            doc.Theme.MinorFonts.ComplexScript = "Andalus";

            Font font = doc.Styles["Normal"].Font;
            Console.WriteLine("Originally the Normal style theme color is: {0} and RGB color is: {1}\n", font.ThemeColor, font.Color);

            // We can use theme font and color instead of default values.
            font.ThemeFont = ThemeFont.Minor;
            font.ThemeColor = ThemeColor.Accent2;

            Assert.That(font.ThemeFont, Is.EqualTo(ThemeFont.Minor));
            Assert.That(font.Name, Is.EqualTo("Algerian"));

            Assert.That(font.ThemeFontAscii, Is.EqualTo(ThemeFont.Minor));
            Assert.That(font.NameAscii, Is.EqualTo("Algerian"));

            Assert.That(font.ThemeFontBi, Is.EqualTo(ThemeFont.Minor));
            Assert.That(font.NameBi, Is.EqualTo("Andalus"));

            Assert.That(font.ThemeFontFarEast, Is.EqualTo(ThemeFont.Minor));
            Assert.That(font.NameFarEast, Is.EqualTo("Aharoni"));

            Assert.That(font.ThemeFontOther, Is.EqualTo(ThemeFont.Minor));
            Assert.That(font.NameOther, Is.EqualTo("Algerian"));

            Assert.That(font.ThemeColor, Is.EqualTo(ThemeColor.Accent2));
            Assert.That(font.Color, Is.EqualTo(Color.Empty));

            // There are several ways of reset them font and color.
            // 1 -  By setting ThemeFont.None/ThemeColor.None:
            font.ThemeFont = ThemeFont.None;
            font.ThemeColor = ThemeColor.None;

            Assert.That(font.ThemeFont, Is.EqualTo(ThemeFont.None));
            Assert.That(font.Name, Is.EqualTo("Algerian"));

            Assert.That(font.ThemeFontAscii, Is.EqualTo(ThemeFont.None));
            Assert.That(font.NameAscii, Is.EqualTo("Algerian"));

            Assert.That(font.ThemeFontBi, Is.EqualTo(ThemeFont.None));
            Assert.That(font.NameBi, Is.EqualTo("Andalus"));

            Assert.That(font.ThemeFontFarEast, Is.EqualTo(ThemeFont.None));
            Assert.That(font.NameFarEast, Is.EqualTo("Aharoni"));

            Assert.That(font.ThemeFontOther, Is.EqualTo(ThemeFont.None));
            Assert.That(font.NameOther, Is.EqualTo("Algerian"));

            Assert.That(font.ThemeColor, Is.EqualTo(ThemeColor.None));
            Assert.That(font.Color, Is.EqualTo(Color.Empty));

            // 2 -  By setting non-theme font/color names:
            font.Name = "Arial";
            font.Color = Color.Blue;

            Assert.That(font.ThemeFont, Is.EqualTo(ThemeFont.None));
            Assert.That(font.Name, Is.EqualTo("Arial"));

            Assert.That(font.ThemeFontAscii, Is.EqualTo(ThemeFont.None));
            Assert.That(font.NameAscii, Is.EqualTo("Arial"));

            Assert.That(font.ThemeFontBi, Is.EqualTo(ThemeFont.None));
            Assert.That(font.NameBi, Is.EqualTo("Arial"));

            Assert.That(font.ThemeFontFarEast, Is.EqualTo(ThemeFont.None));
            Assert.That(font.NameFarEast, Is.EqualTo("Arial"));

            Assert.That(font.ThemeFontOther, Is.EqualTo(ThemeFont.None));
            Assert.That(font.NameOther, Is.EqualTo("Arial"));

            Assert.That(font.ThemeColor, Is.EqualTo(ThemeColor.None));
            Assert.That(font.Color.ToArgb(), Is.EqualTo(Color.Blue.ToArgb()));
            //ExEnd
        }

        [Test]
        public void CreateThemedStyle()
        {
            //ExStart
            //ExFor:Font.ThemeFont
            //ExFor:Font.ThemeColor
            //ExFor:Font.TintAndShade
            //ExFor:ThemeFont
            //ExFor:ThemeColor
            //ExSummary:Shows how to create and use themed style.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln();

            // Create some style with theme font properties.
            Style style = doc.Styles.Add(StyleType.Paragraph, "ThemedStyle");
            style.Font.ThemeFont = ThemeFont.Major;
            style.Font.ThemeColor = ThemeColor.Accent5;
            style.Font.TintAndShade = 0.3;

            builder.ParagraphFormat.StyleName = "ThemedStyle";
            builder.Writeln("Text with themed style");
            //ExEnd

            Run run = (Run)((Paragraph)builder.CurrentParagraph.PreviousSibling).FirstChild;

            Assert.That(run.Font.ThemeFont, Is.EqualTo(ThemeFont.Major));
            Assert.That(run.Font.Name, Is.EqualTo("Times New Roman"));

            Assert.That(run.Font.ThemeFontAscii, Is.EqualTo(ThemeFont.Major));
            Assert.That(run.Font.NameAscii, Is.EqualTo("Times New Roman"));

            Assert.That(run.Font.ThemeFontBi, Is.EqualTo(ThemeFont.Major));
            Assert.That(run.Font.NameBi, Is.EqualTo("Times New Roman"));

            Assert.That(run.Font.ThemeFontFarEast, Is.EqualTo(ThemeFont.Major));
            Assert.That(run.Font.NameFarEast, Is.EqualTo("Times New Roman"));

            Assert.That(run.Font.ThemeFontOther, Is.EqualTo(ThemeFont.Major));
            Assert.That(run.Font.NameOther, Is.EqualTo("Times New Roman"));

            Assert.That(run.Font.ThemeColor, Is.EqualTo(ThemeColor.Accent5));
            Assert.That(run.Font.Color, Is.EqualTo(Color.Empty));
        }

        [Test]
        public void FontInfoEmbeddingLicensingRights()
        {
            //ExStart:FontInfoEmbeddingLicensingRights
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:FontInfo.EmbeddingLicensingRights
            //ExFor:FontEmbeddingUsagePermissions
            //ExFor:FontEmbeddingLicensingRights
            //ExFor:FontEmbeddingLicensingRights.EmbeddingUsagePermissions
            //ExFor:FontEmbeddingLicensingRights.BitmapEmbeddingOnly
            //ExFor:FontEmbeddingLicensingRights.NoSubsetting
            //ExSummary:Shows how to get license rights information for embedded fonts (FontInfo).
            Document doc = new Document(MyDir + "Embedded font rights.docx");

            // Get the list of document fonts.
            FontInfoCollection fontInfos = doc.FontInfos;
            foreach (FontInfo fontInfo in fontInfos) 
            {
                if (fontInfo.EmbeddingLicensingRights != null)
                {
                    Console.WriteLine(fontInfo.EmbeddingLicensingRights.EmbeddingUsagePermissions);
                    Console.WriteLine(fontInfo.EmbeddingLicensingRights.BitmapEmbeddingOnly);
                    Console.WriteLine(fontInfo.EmbeddingLicensingRights.NoSubsetting);
                }
            }
            //ExEnd:FontInfoEmbeddingLicensingRights
        }

        [Test]
        public void PhysicalFontInfoEmbeddingLicensingRights()
        {
            //ExStart:PhysicalFontInfoEmbeddingLicensingRights
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:PhysicalFontInfo.EmbeddingLicensingRights
            //ExSummary:Shows how to get license rights information for embedded fonts (PhysicalFontInfo).
            FontSettings settings = FontSettings.DefaultInstance;
            FontSourceBase source = settings.GetFontsSources()[0];

            // Get the list of available fonts.
            IList<PhysicalFontInfo> fontInfos = source.GetAvailableFonts();
            foreach (PhysicalFontInfo fontInfo in fontInfos)
            {
                if (fontInfo.EmbeddingLicensingRights != null)
                {
                    Console.WriteLine(fontInfo.EmbeddingLicensingRights.EmbeddingUsagePermissions);
                    Console.WriteLine(fontInfo.EmbeddingLicensingRights.BitmapEmbeddingOnly);
                    Console.WriteLine(fontInfo.EmbeddingLicensingRights.NoSubsetting);
                }
            }
            //ExEnd:PhysicalFontInfoEmbeddingLicensingRights
        }

        [Test]
        public void NumberSpacing()
        {
            //ExStart:NumberSpacing
            //GistId:95fdae949cefbf2ce485acc95cccc495
            //ExFor:Font.NumberSpacing
            //ExFor:NumSpacing
            //ExSummary:Shows how to set the spacing type of the numeral.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // This effect is only supported in newer versions of MS Word.
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2019);

            builder.Write("1 ");
            builder.Write("This is an example");

            Run run = doc.FirstSection.Body.FirstParagraph.Runs[0];
            if (run.Font.NumberSpacing == NumSpacing.Default)
                run.Font.NumberSpacing = NumSpacing.Proportional;

            doc.Save(ArtifactsDir + "Fonts.NumberSpacing.docx");
            //ExEnd:NumberSpacing

            doc = new Document(ArtifactsDir + "Fonts.NumberSpacing.docx");

            run = doc.FirstSection.Body.FirstParagraph.Runs[0];
            Assert.That(run.Font.NumberSpacing, Is.EqualTo(NumSpacing.Proportional));
        }
    }
}