// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#if NET462 || NETCOREAPP2_1 || JAVA
using System;
using System.Collections;
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

            Assert.AreEqual("Hello world!", run.GetText().Trim());
            Assert.AreEqual("Courier New", run.Font.Name);
            Assert.AreEqual(36, run.Font.Size);
            Assert.AreEqual(Color.Yellow.ToArgb(), run.Font.HighlightColor.ToArgb());
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

            Assert.AreEqual("all capitals", run.GetText().Trim());
            Assert.True(run.Font.AllCaps);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("Small Capitals", run.GetText().Trim());
            Assert.True(run.Font.SmallCaps);
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
            Assert.AreEqual(5, allFonts.Count); //ExSkip

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

            Assert.IsFalse(doc.FontInfos.EmbedTrueTypeFonts);
            Assert.IsFalse(doc.FontInfos.EmbedSystemFonts);
            Assert.IsFalse(doc.FontInfos.SaveSubsetFonts);
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

            if (embedAllFonts)
                Assert.That(25000, Is.LessThan(new FileInfo(ArtifactsDir + "Font.FontInfoCollection.docx").Length));
            else
                Assert.That(15000, Is.AtLeast(new FileInfo(ArtifactsDir + "Font.FontInfoCollection.docx").Length));
            //ExEnd
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

            Assert.AreEqual("Text with a single-line strikethrough.", run.GetText().Trim());
            Assert.True(run.Font.StrikeThrough);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("Text with a double-line strikethrough.", run.GetText().Trim());
            Assert.True(run.Font.DoubleStrikeThrough);
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
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

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

            Assert.AreEqual("Raised text.", run.GetText().Trim());
            Assert.AreEqual(5, run.Font.Position);

            doc = new Document(ArtifactsDir + "Font.PositionSubscript.docx");
            run = doc.FirstSection.Body.FirstParagraph.Runs[1];

            Assert.AreEqual("Lowered text.", run.GetText().Trim());
            Assert.AreEqual(-10, run.Font.Position);

            run = doc.FirstSection.Body.FirstParagraph.Runs[3];

            Assert.AreEqual("Subscript.", run.GetText().Trim());
            Assert.True(run.Font.Subscript);

            run = doc.FirstSection.Body.FirstParagraph.Runs[4];

            Assert.AreEqual("Superscript.", run.GetText().Trim());
            Assert.True(run.Font.Superscript);
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

            Assert.AreEqual("Wide characters", run.GetText().Trim());
            Assert.AreEqual(150, run.Font.Scaling);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("Expanded by 1pt", run.GetText().Trim());
            Assert.AreEqual(1, run.Font.Spacing);

            run = doc.FirstSection.Body.Paragraphs[2].Runs[0];

            Assert.AreEqual("Condensed by 1pt", run.GetText().Trim());
            Assert.AreEqual(-1, run.Font.Spacing);
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

            Assert.AreEqual("Hello world!", run.GetText().Trim());
            Assert.True(run.Font.Italic);
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

            Assert.AreEqual("This text is engraved.", run.GetText().Trim());
            Assert.True(run.Font.Engrave);
            Assert.False(run.Font.Emboss);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("This text is embossed.", run.GetText().Trim());
            Assert.False(run.Font.Engrave);
            Assert.True(run.Font.Emboss);
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

            Assert.AreEqual("This text has a shadow.", run.GetText().Trim());
            Assert.True(run.Font.Shadow);
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

            Assert.AreEqual("This text has an outline.", run.GetText().Trim());
            Assert.True(run.Font.Outline);
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

            Assert.AreEqual("This text will not be visible in the document.", run.GetText().Trim());
            Assert.True(run.Font.Hidden);
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

            Assert.AreEqual("TALLY. (Kerning not applied)", run.GetText().Trim());
            Assert.AreEqual(24, run.Font.Kerning);
            Assert.AreEqual(18, run.Font.Size);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("TALLY. (Kerning applied)", run.GetText().Trim());
            Assert.AreEqual(12, run.Font.Kerning);
            Assert.AreEqual(18, run.Font.Size);
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

            Assert.AreEqual("Proofing has been disabled, so these spelking errrs will not display red lines underneath.", run.GetText().Trim());
            Assert.True(run.Font.NoProofing);
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

            Assert.AreEqual("Привет!", run.GetText().Trim());
            Assert.AreEqual(1033, run.Font.LocaleId);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("Привет!", run.GetText().Trim());
            Assert.AreEqual(1049, run.Font.LocaleId);
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

            Assert.AreEqual("Underlined text.", run.GetText().Trim());
            Assert.AreEqual(Underline.Dotted, run.Font.Underline);
            Assert.AreEqual(Color.Red.ToArgb(), run.Font.UnderlineColor.ToArgb());
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

            Assert.AreEqual("Text treated as complex script.", run.GetText().Trim());
            Assert.True(run.Font.ComplexScript);
        }

        [Test]
        public void SparklingText()
        {
            //ExStart
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

            Assert.AreEqual("Text with a sparkle effect.", run.GetText().Trim());
            Assert.AreEqual(TextEffect.SparkleText, run.Font.TextEffect);
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

            Assert.AreEqual("White text on an orange background with a two-tone texture.", run.GetText().Trim());
            Assert.AreEqual(Color.White.ToArgb(), run.Font.Color.ToArgb());

            Assert.AreEqual(TextureIndex.TextureDiagonalUp, run.Font.Shading.Texture);
            Assert.AreEqual(Color.OrangeRed.ToArgb(), run.Font.Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.DarkBlue.ToArgb(), run.Font.Shading.ForegroundPatternColor.ToArgb());
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
                        Assert.AreEqual("مرحبًا", run.GetText().Trim());
                        Assert.True(run.Font.Bidi);
                        break;
                    case 1:
                        Assert.AreEqual("Hello world!", run.GetText().Trim());
                        Assert.False(run.Font.Bidi);
                        break;
                }

                Assert.AreEqual(1033, run.Font.LocaleId);
                Assert.AreEqual(16, run.Font.Size);
                Assert.AreEqual("Courier New", run.Font.Name);
                Assert.False(run.Font.Italic);
                Assert.False(run.Font.Bold);
                Assert.AreEqual(1025, run.Font.LocaleIdBi);
                Assert.AreEqual(24, run.Font.SizeBi);
                Assert.AreEqual("Andalus", run.Font.NameBi);
                Assert.True(run.Font.ItalicBi);
                Assert.True(run.Font.BoldBi);
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

            Assert.AreEqual("Hello world!", run.GetText().Trim());
            Assert.AreEqual(1033, run.Font.LocaleId);
            Assert.AreEqual("Courier New", run.Font.Name);
            Assert.AreEqual(2052, run.Font.LocaleIdFarEast);
            Assert.AreEqual("SimSun", run.Font.NameFarEast);

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("你好世界", run.GetText().Trim());
            Assert.AreEqual(1033, run.Font.LocaleId);
            Assert.AreEqual("SimSun", run.Font.Name);
            Assert.AreEqual(2052, run.Font.LocaleIdFarEast);
            Assert.AreEqual("SimSun", run.Font.NameFarEast);
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
            Assert.AreEqual("Calibri", builder.Font.Name);

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

            Assert.AreEqual("Hello, Привет", run.GetText().Trim());
            Assert.AreEqual("Calibri", run.Font.Name);
            Assert.AreEqual("Calibri", run.Font.NameAscii);
            Assert.AreEqual("Courier New", run.Font.NameOther);
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
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true).OfType<Run>())
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

            Assert.AreEqual("Text originally in \"Emphasis\" style", docRun.GetText().Trim());
            Assert.AreEqual(StyleIdentifier.Strong, docRun.Font.StyleIdentifier);
            Assert.AreEqual("Strong", docRun.Font.StyleName);

            docRun = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("Text originally in \"Intense Emphasis\" style", docRun.GetText().Trim());
            Assert.AreEqual(StyleIdentifier.Strong, docRun.Font.StyleIdentifier);
            Assert.AreEqual("Strong", docRun.Font.StyleName);
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

            Assert.True(style.BuiltIn);

            // Create a custom style and add it to the collection.
            // Custom styles such as this will have the "BuiltIn" flag set to "false". 
            style = doc.Styles.Add(StyleType.Character, "MyStyle");
            style.Font.Color = Color.Navy;
            style.Font.Name = "Courier New";

            Assert.False(style.BuiltIn);
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
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true).OfType<Run>())
            {
                Style charStyle = run.Font.Style;

                if (!charStyle.BuiltIn)
                    run.Font.Underline = Underline.Double;
            }

            doc.Save(ArtifactsDir + "Font.Style.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Style.docx");
            Run docRun = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("This text is in a custom style.", docRun.GetText().Trim());
            Assert.AreEqual("MyStyle", docRun.Font.StyleName);
            Assert.False(docRun.Font.Style.BuiltIn);
            Assert.AreEqual(Underline.Double, docRun.Font.Underline);
        }
        
        [Test]
        public void GetAvailableFonts()
        {
            //ExStart
            //ExFor:Fonts.PhysicalFontInfo
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

            Assert.AreEqual(folderFontSource[0].GetAvailableFonts().Count, 
                Directory.EnumerateFiles(FontsDir, "*.*", SearchOption.AllDirectories).Count(f => f.EndsWith(".ttf") || f.EndsWith(".otf")));
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
            Assert.AreEqual(Color.Empty.ToArgb(), builder.Font.Color.ToArgb());

            // The default color for text is black. If the color of the background is dark, black text will be difficult to see.
            // To solve this problem, the AutoColor property will display this text in white.
            builder.Font.Shading.BackgroundPatternColor = Color.DarkBlue;

            builder.Writeln("The text color automatically chosen for this run is white.");

            Assert.AreEqual(Color.White.ToArgb(), doc.FirstSection.Body.Paragraphs[0].Runs[0].Font.AutoColor.ToArgb());

            // If we change the background to a light color, black will be a more
            // suitable text color than white so that the auto color will display it in black.
            builder.Font.Shading.BackgroundPatternColor = Color.LightBlue;

            builder.Writeln("The text color automatically chosen for this run is black.");

            Assert.AreEqual(Color.Black.ToArgb(), doc.FirstSection.Body.Paragraphs[1].Runs[0].Font.AutoColor.ToArgb());

            doc.Save(ArtifactsDir + "Font.SetFontAutoColor.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.SetFontAutoColor.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("The text color automatically chosen for this run is white.", run.GetText().Trim());
            Assert.AreEqual(Color.Empty.ToArgb(), run.Font.Color.ToArgb());
            Assert.AreEqual(Color.DarkBlue.ToArgb(), run.Font.Shading.BackgroundPatternColor.ToArgb());

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("The text color automatically chosen for this run is black.", run.GetText().Trim());
            Assert.AreEqual(Color.Empty.ToArgb(), run.Font.Color.ToArgb());
            Assert.AreEqual(Color.LightBlue.ToArgb(), run.Font.Shading.BackgroundPatternColor.ToArgb());
        }

        //ExStart
        //ExFor:Font.Hidden
        //ExFor:Paragraph.Accept
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
        //ExFor:Node.Accept
        //ExFor:Paragraph.ParagraphBreakFont
        //ExFor:Table.Accept
        //ExSummary:Shows how to use a DocumentVisitor implementation to remove all hidden content from a document.
        [Test] //ExSkip
        public void RemoveHiddenContentFromDocument()
        {
            Document doc = new Document(MyDir + "Hidden content.docx");
            Assert.AreEqual(26, doc.GetChildNodes(NodeType.Paragraph, true).Count); //ExSkip
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count); //ExSkip

            RemoveHiddenContentVisitor hiddenContentRemover = new RemoveHiddenContentVisitor();

            // Below are three types of fields which can accept a document visitor,
            // which will allow it to visit the accepting node, and then traverse its child nodes in a depth-first manner.
            // 1 -  Paragraph node:
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 4, true);
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
            Assert.AreEqual(20, doc.GetChildNodes(NodeType.Paragraph, true).Count); //ExSkip
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, true).Count); //ExSkip

            foreach (Node node in doc.GetChildNodes(NodeType.Any, true))
            {
                switch (node)
                {
                    case FieldStart fieldStart:
                        Assert.False(fieldStart.Font.Hidden);
                        break;
                    case FieldEnd fieldEnd:
                        Assert.False(fieldEnd.Font.Hidden);
                        break;
                    case FieldSeparator fieldSeparator:
                        Assert.False(fieldSeparator.Font.Hidden);
                        break;
                    case Run run:
                        Assert.False(run.Font.Hidden);
                        break;
                    case Paragraph paragraph:
                        Assert.False(paragraph.ParagraphBreakFont.Hidden);
                        break;
                    case FormField formField:
                        Assert.False(formField.Font.Hidden);
                        break;
                    case GroupShape groupShape:
                        Assert.False(groupShape.Font.Hidden);
                        break;
                    case Shape shape:
                        Assert.False(shape.Font.Hidden);
                        break;
                    case Comment comment:
                        Assert.False(comment.Font.Hidden);
                        break;
                    case Footnote footnote:
                        Assert.False(footnote.Font.Hidden);
                        break;
                    case SpecialChar specialChar:
                        Assert.False(specialChar.Font.Hidden);
                        break;
                }
            } 
        }

        [Test]
        public void DefaultFonts()
        {
            //ExStart
            //ExFor:Fonts.FontInfoCollection.Contains(String)
            //ExFor:Fonts.FontInfoCollection.Count
            //ExSummary:Shows info about the fonts that are present in the blank document.
            Document doc = new Document();

            // A blank document contains 3 default fonts. Each font in the document
            // will have a corresponding FontInfo object which contains details about that font.
            Assert.AreEqual(3, doc.FontInfos.Count);

            Assert.True(doc.FontInfos.Contains("Times New Roman"));
            Assert.AreEqual(204, doc.FontInfos["Times New Roman"].Charset);

            Assert.True(doc.FontInfos.Contains("Symbol"));
            Assert.True(doc.FontInfos.Contains("Arial"));
            //ExEnd
        }

        [Test]
        public void ExtractEmbeddedFont()
        {
            //ExStart
            //ExFor:Fonts.EmbeddedFontFormat
            //ExFor:Fonts.EmbeddedFontStyle
            //ExFor:Fonts.FontInfo.GetEmbeddedFont(EmbeddedFontFormat,EmbeddedFontStyle)
            //ExFor:Fonts.FontInfo.GetEmbeddedFontAsOpenType(EmbeddedFontStyle)
            //ExFor:Fonts.FontInfoCollection.Item(Int32)
            //ExFor:Fonts.FontInfoCollection.Item(String)
            //ExSummary:Shows how to extract an embedded font from a document, and save it to the local file system.
            Document doc = new Document(MyDir + "Embedded font.docx");

            FontInfo embeddedFont = doc.FontInfos["Alte DIN 1451 Mittelschrift"];
            byte[] embeddedFontBytes = embeddedFont.GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular);
            Assert.IsNotNull(embeddedFontBytes); //ExSkip

            File.WriteAllBytes(ArtifactsDir + "Alte DIN 1451 Mittelschrift.ttf", embeddedFontBytes);
            
            // Embedded font formats may be different in other formats such as .doc.
            // We need to know the correct format before we can extract the font.
            doc = new Document(MyDir + "Embedded font.doc");

            Assert.IsNull(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular));
            Assert.IsNotNull(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFont(EmbeddedFontFormat.EmbeddedOpenType, EmbeddedFontStyle.Regular));

            // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType.
            embeddedFontBytes = doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFontAsOpenType(EmbeddedFontStyle.Regular);

            File.WriteAllBytes(ArtifactsDir + "Alte DIN 1451 Mittelschrift.otf", embeddedFontBytes);
            //ExEnd
        }

        [Test]
        public void GetFontInfoFromFile() 
        {
            //ExStart
            //ExFor:Fonts.FontFamily
            //ExFor:Fonts.FontPitch
            //ExFor:Fonts.FontInfo.AltName
            //ExFor:Fonts.FontInfo.Charset
            //ExFor:Fonts.FontInfo.Family
            //ExFor:Fonts.FontInfo.Panose
            //ExFor:Fonts.FontInfo.Pitch
            //ExFor:Fonts.FontInfoCollection.GetEnumerator
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

            Assert.AreEqual(new[] { 2, 15, 5, 2, 2, 2, 4, 3, 2, 4 }, doc.FontInfos["Calibri"].Panose);
            Assert.AreEqual(new[] { 2, 15, 3, 2, 2, 2, 4, 3, 2, 4 }, doc.FontInfos["Calibri Light"].Panose);
            Assert.AreEqual(new[] { 2, 2, 6, 3, 5, 4, 5, 2, 3, 4 }, doc.FontInfos["Times New Roman"].Panose);
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
            Assert.AreEqual(14.6484375d, builder.Font.LineSpacing);

            builder.Font.Name = "Times New Roman";
            Assert.AreEqual(13.798828125d, builder.Font.LineSpacing);
            //ExEnd
        }

        [Test]
        public void HasDmlEffect()
        {
            //ExStart
            //ExFor:Font.HasDmlEffect(TextDmlEffect)
            //ExSummary:Shows how to check if a run displays a DrawingML text effect.
            Document doc = new Document(MyDir + "DrawingML text effects.docx");
            
            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;
            
            Assert.True(runs[0].Font.HasDmlEffect(TextDmlEffect.Shadow));
            Assert.True(runs[1].Font.HasDmlEffect(TextDmlEffect.Shadow));
            Assert.True(runs[2].Font.HasDmlEffect(TextDmlEffect.Reflection));
            Assert.True(runs[3].Font.HasDmlEffect(TextDmlEffect.Effect3D));
            Assert.True(runs[4].Font.HasDmlEffect(TextDmlEffect.Fill));
            //ExEnd
        }
        
        [Test, Category("IgnoreOnJenkins")]
        public void CheckScanUserFontsFolder()
        {
            // On Windows 10 fonts may be installed either into system folder "%windir%\fonts" for all users
            // or into user folder "%userprofile%\AppData\Local\Microsoft\Windows\Fonts" for current user.
            SystemFontSource systemFontSource = new SystemFontSource();
            Assert.NotNull(systemFontSource.GetAvailableFonts()
                    .FirstOrDefault(x => x.FilePath.Contains("\\AppData\\Local\\Microsoft\\Windows\\Fonts")),
                "Fonts did not install to the user font folder");
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
            
            Assert.AreEqual(ThemeFont.Minor, font.ThemeFont);
            Assert.AreEqual("Algerian", font.Name);
            
            Assert.AreEqual(ThemeFont.Minor, font.ThemeFontAscii);
            Assert.AreEqual("Algerian", font.NameAscii);

            Assert.AreEqual(ThemeFont.Minor, font.ThemeFontBi);
            Assert.AreEqual("Andalus", font.NameBi);

            Assert.AreEqual(ThemeFont.Minor, font.ThemeFontFarEast);
            Assert.AreEqual("Aharoni", font.NameFarEast);

            Assert.AreEqual(ThemeFont.Minor, font.ThemeFontOther);
            Assert.AreEqual("Algerian", font.NameOther);

            Assert.AreEqual(ThemeColor.Accent2, font.ThemeColor);
            Assert.AreEqual(Color.Empty, font.Color);

            // There are several ways of reset them font and color.
            // 1 -  By setting ThemeFont.None/ThemeColor.None:
            font.ThemeFont = ThemeFont.None;
            font.ThemeColor = ThemeColor.None;

            Assert.AreEqual(ThemeFont.None, font.ThemeFont);
            Assert.AreEqual("Algerian", font.Name);

            Assert.AreEqual(ThemeFont.None, font.ThemeFontAscii);
            Assert.AreEqual("Algerian", font.NameAscii);

            Assert.AreEqual(ThemeFont.None, font.ThemeFontBi);
            Assert.AreEqual("Andalus", font.NameBi);

            Assert.AreEqual(ThemeFont.None, font.ThemeFontFarEast);
            Assert.AreEqual("Aharoni", font.NameFarEast);

            Assert.AreEqual(ThemeFont.None, font.ThemeFontOther);
            Assert.AreEqual("Algerian", font.NameOther);

            Assert.AreEqual(ThemeColor.None, font.ThemeColor);
            Assert.AreEqual(Color.Empty, font.Color);

            // 2 -  By setting non-theme font/color names:
            font.Name = "Arial";
            font.Color = Color.Blue;

            Assert.AreEqual(ThemeFont.None, font.ThemeFont);
            Assert.AreEqual("Arial", font.Name);

            Assert.AreEqual(ThemeFont.None, font.ThemeFontAscii);
            Assert.AreEqual("Arial", font.NameAscii);

            Assert.AreEqual(ThemeFont.None, font.ThemeFontBi);
            Assert.AreEqual("Arial", font.NameBi);

            Assert.AreEqual(ThemeFont.None, font.ThemeFontFarEast);
            Assert.AreEqual("Arial", font.NameFarEast);

            Assert.AreEqual(ThemeFont.None, font.ThemeFontOther);
            Assert.AreEqual("Arial", font.NameOther);

            Assert.AreEqual(ThemeColor.None, font.ThemeColor);
            Assert.AreEqual(Color.Blue.ToArgb(), font.Color.ToArgb());
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

            Assert.AreEqual(ThemeFont.Major, run.Font.ThemeFont);
            Assert.AreEqual("Times New Roman", run.Font.Name);

            Assert.AreEqual(ThemeFont.Major, run.Font.ThemeFontAscii);
            Assert.AreEqual("Times New Roman", run.Font.NameAscii);

            Assert.AreEqual(ThemeFont.Major, run.Font.ThemeFontBi);
            Assert.AreEqual("Times New Roman", run.Font.NameBi);

            Assert.AreEqual(ThemeFont.Major, run.Font.ThemeFontFarEast);
            Assert.AreEqual("Times New Roman", run.Font.NameFarEast);

            Assert.AreEqual(ThemeFont.Major, run.Font.ThemeFontOther);
            Assert.AreEqual("Times New Roman", run.Font.NameOther);

            Assert.AreEqual(ThemeColor.Accent5, run.Font.ThemeColor);
            Assert.AreEqual(Color.Empty, run.Font.Color);
        }
    }
}
#endif