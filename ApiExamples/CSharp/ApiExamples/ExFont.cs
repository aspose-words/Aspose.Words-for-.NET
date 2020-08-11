// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
using System.Text.RegularExpressions;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Fonts;
using Aspose.Words.Tables;
using NUnit.Framework;

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
            //ExSummary:Shows how to add a formatted run of text to a document using the object model.
            Document doc = new Document();

            // Create a new run of text
            Run run = new Run(doc, "Hello");

            // Specify character formatting for the run of text
            Aspose.Words.Font f = run.Font;
            f.Name = "Courier New";
            f.Size = 36;
            f.HighlightColor = Color.Yellow;

            // Append the run of text to the end of the first paragraph
            // in the body of the first section of the document
            doc.FirstSection.Body.FirstParagraph.AppendChild(run);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            run = doc.FirstSection.Body.FirstParagraph.Runs[0];

            Assert.AreEqual("Hello", run.GetText().Trim());
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
            //ExSummary:Shows how to use all capitals and small capitals character formatting properties.
            Document doc = new Document();
            Paragraph para = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            Run run = new Run(doc, "All capitals");
            run.Font.AllCaps = true;
            para.AppendChild(run);

            run = new Run(doc, "SMALL CAPITALS");
            run.Font.SmallCaps = true;
            para.AppendChild(run);

            doc.Save(ArtifactsDir + "Font.Caps.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Caps.docx");
            run = doc.FirstSection.Body.FirstParagraph.Runs[0];

            Assert.AreEqual("All capitals", run.GetText().Trim());
            Assert.True(run.Font.AllCaps);

            run = doc.FirstSection.Body.FirstParagraph.Runs[1];

            Assert.AreEqual("SMALL CAPITALS", run.GetText().Trim());
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

            FontInfoCollection fonts = doc.FontInfos;
            Assert.AreEqual(5, fonts.Count); //ExSkip

            // The fonts info extracted from this document does not necessarily mean that the fonts themselves are
            // used in the document. If a font is present but not used then most likely they were referenced at some time
            // and then removed from the Document
            for (int i = 0; i < fonts.Count; i++)
            {
                Console.WriteLine($"Font index #{i}");
                Console.WriteLine($"\tName: {fonts[i].Name}");
                Console.WriteLine($"\tIs {(fonts[i].IsTrueType ? "" : "not ")}a trueType font");
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

        [Test]
        public void FontInfoCollection()
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
            fontInfos.EmbedTrueTypeFonts = true;
            fontInfos.EmbedSystemFonts = false;
            fontInfos.SaveSubsetFonts = false;

            doc.Save(ArtifactsDir + "Font.FontInfoCollection.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.FontInfoCollection.docx");
            fontInfos = doc.FontInfos;

            Assert.True(fontInfos.EmbedTrueTypeFonts);
            Assert.False(fontInfos.EmbedSystemFonts);
            Assert.False(fontInfos.SaveSubsetFonts);
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
            //ExSummary:Shows how to use strike-through character formatting properties.
            Document doc = new Document();
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

            Run run = new Run(doc, "Double strike through text");
            run.Font.DoubleStrikeThrough = true;
            para.AppendChild(run);

            run = new Run(doc, "Single strike through text");
            run.Font.StrikeThrough = true;
            para.AppendChild(run);

            doc.Save(ArtifactsDir + "Font.StrikeThrough.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.StrikeThrough.docx");
            run = doc.FirstSection.Body.FirstParagraph.Runs[0];

            Assert.AreEqual("Double strike through text", run.GetText().Trim());
            Assert.True(run.Font.DoubleStrikeThrough);

            run = doc.FirstSection.Body.FirstParagraph.Runs[1];

            Assert.AreEqual("Single strike through text", run.GetText().Trim());
            Assert.True(run.Font.StrikeThrough);
        }

        [Test]
        public void PositionSubscript()
        {
            //ExStart
            //ExFor:Font.Position
            //ExFor:Font.Subscript
            //ExFor:Font.Superscript
            //ExSummary:Shows how to use subscript, superscript, complex script, text effects, and baseline text position properties.
            Document doc = new Document();
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

            // Add a run of text that is raised 5 points above the baseline
            Run run = new Run(doc, "Raised text");
            run.Font.Position = 5;
            para.AppendChild(run);

            // Add a run of normal text
            run = new Run(doc, "Normal text");
            para.AppendChild(run);

            // Add a run of text that appears as subscript
            run = new Run(doc, "Subscript");
            run.Font.Subscript = true;
            para.AppendChild(run);

            // Add a run of text that appears as superscript
            run = new Run(doc, "Superscript");
            run.Font.Superscript = true;
            para.AppendChild(run);

            doc.Save(ArtifactsDir + "Font.PositionSubscript.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.PositionSubscript.docx");
            run = doc.FirstSection.Body.FirstParagraph.Runs[0];

            Assert.AreEqual("Raised text", run.GetText().Trim());
            Assert.AreEqual(5, run.Font.Position);

            run = doc.FirstSection.Body.FirstParagraph.Runs[2];

            Assert.AreEqual("Subscript", run.GetText().Trim());
            Assert.True(run.Font.Subscript);

            run = doc.FirstSection.Body.FirstParagraph.Runs[3];

            Assert.AreEqual("Superscript", run.GetText().Trim());
            Assert.True(run.Font.Superscript);
        }

        [Test]
        public void ScalingSpacing()
        {
            //ExStart
            //ExFor:Font.Scaling
            //ExFor:Font.Spacing
            //ExSummary:Shows how to use character scaling and spacing properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a run of text with characters 150% width of normal characters
            builder.Font.Scaling = 150;
            builder.Writeln("Wide characters");

            // Add a run of text with extra 1pt space between characters
            builder.Font.Spacing = 1;
            builder.Writeln("Expanded by 1pt");

            // Add a run of text with space between characters reduced by 1pt
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
            //ExSummary:Shows how to italicize a run of text.
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
            //ExSummary:Shows the difference between embossing and engraving text via font formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Size = 36;
            builder.Font.Color = Color.White;
            builder.Font.Engrave = true;

            builder.Writeln("This text is engraved.");

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
            builder.Font.Size = 36;
            builder.Font.Shadow = true;

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
            builder.Font.Size = 36;
            builder.Font.Outline = true;

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
            //ExSummary:Shows how to create a hidden run of text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Size = 36;
            builder.Font.Hidden = true;

            // With the Hidden flag set to true, we can add text that will be present but invisible in the document
            // It is not recommended to use this as a way of hiding sensitive information as the text is still easily reachable
            builder.Writeln("This text won't be visible in the document.");

            doc.Save(ArtifactsDir + "Font.Hidden.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Hidden.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("This text won't be visible in the document.", run.GetText().Trim());
            Assert.True(run.Font.Hidden);
        }

        [Test]
        public void Kerning()
        {
            //ExStart
            //ExFor:Font.Kerning
            //ExSummary:Shows how to specify the font size at which kerning starts.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Arial Black";

            // Set the font's kerning size threshold and font size 
            builder.Font.Kerning = 24;
            builder.Font.Size = 18;

            // The font size falls below the kerning threshold so kerning will not be applied
            builder.Writeln("TALLY. (Kerning not applied)");

            // If we add runs of text with a document builder's writing methods,
            // the Font attributes of any new runs will inherit the values from the Font attributes of the previous runs
            // The font size is still 18, and we will change the kerning threshold to a value below that
            builder.Font.Kerning = 12;
            
            // Kerning has now been applied to this run
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
            //ExSummary:Shows how to specify that the run of text is not to be spell checked by Microsoft Word.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.NoProofing = true;

            builder.Writeln("Proofing has been disabled for this run, so these spelking errrs will not display red lines underneath.");

            doc.Save(ArtifactsDir + "Font.NoProofing.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.NoProofing.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("Proofing has been disabled for this run, so these spelking errrs will not display red lines underneath.", run.GetText().Trim());
            Assert.True(run.Font.NoProofing);
        }

        [Test]
        public void LocaleId()
        {
            //ExStart
            //ExFor:Font.LocaleId
            //ExSummary:Shows how to specify the language of a text run so Microsoft Word can use a proper spell checker.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify the locale so Microsoft Word recognizes this text as Russian
            builder.Font.LocaleId = new CultureInfo("ru-RU", false).LCID;
            builder.Writeln("Привет!");

            doc.Save(ArtifactsDir + "Font.LocaleId.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.LocaleId.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("Привет!", run.GetText().Trim());
            Assert.AreEqual(1049, run.Font.LocaleId);
        }

        [Test]
        public void Underlines()
        {
            //ExStart
            //ExFor:Font.Underline
            //ExFor:Font.UnderlineColor
            //ExSummary:Shows how use the underline character formatting properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set an underline color and style
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
            //ExSummary:Shows how to make a run that's always treated as complex script.
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
            
            // Font animation effects are only visible in older versions of Microsoft Word
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
            //ExSummary:Shows how to apply shading for a run of text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shading shd = builder.Font.Shading;
            shd.Texture = TextureIndex.TextureDiagonalUp;
            shd.BackgroundPatternColor = Color.OrangeRed;
            shd.ForegroundPatternColor = Color.DarkBlue;

            builder.Font.Color = Color.White;

            builder.Writeln("White text on an orange background with texture.");

            doc.Save(ArtifactsDir + "Font.Shading.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Shading.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("White text on an orange background with texture.", run.GetText().Trim());
            Assert.AreEqual(Color.White.ToArgb(), run.Font.Color.ToArgb());

            Assert.AreEqual(TextureIndex.TextureDiagonalUp, run.Font.Shading.Texture);
            Assert.AreEqual(Color.OrangeRed.ToArgb(), run.Font.Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.DarkBlue.ToArgb(), run.Font.Shading.ForegroundPatternColor.ToArgb());
        }

        [Test]
        public void Bidi()
        {
            //ExStart
            //ExFor:Font.Bidi
            //ExFor:Font.NameBi
            //ExFor:Font.SizeBi
            //ExFor:Font.ItalicBi
            //ExFor:Font.BoldBi
            //ExFor:Font.LocaleIdBi
            //ExSummary:Shows how to insert and format right-to-left text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Signal to Microsoft Word that this run of text contains right-to-left text
            builder.Font.Bidi = true;

            // Specify the font and font size to be used for the right-to-left text
            builder.Font.NameBi = "Andalus";
            builder.Font.SizeBi = 48;

            // Specify that the right-to-left text in this run is bold and italic
            builder.Font.ItalicBi = true;
            builder.Font.BoldBi = true;

            // Specify the locale so Microsoft Word recognizes this text as Arabic - Saudi Arabia
            builder.Font.LocaleIdBi = new CultureInfo("ar-AR", false).LCID;

            // Insert some Arabic text
            builder.Writeln("مرحبًا");

            doc.Save(ArtifactsDir + "Font.Bidi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Bidi.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("مرحبًا", run.GetText().Trim());
            Assert.AreEqual(1033, run.Font.LocaleId);
            Assert.True(run.Font.Bidi);
            Assert.AreEqual(48, run.Font.SizeBi);
            Assert.AreEqual("Andalus", run.Font.NameBi);
            Assert.True(run.Font.ItalicBi);
            Assert.True(run.Font.BoldBi);
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

            // Specify the font name
            builder.Font.NameFarEast = "SimSun";

            // Specify the locale so Microsoft Word recognizes this text as Chinese
            builder.Font.LocaleIdFarEast = new CultureInfo("zh-CN", false).LCID;

            // Insert some Chinese text
            builder.Writeln("你好世界");

            doc.Save(ArtifactsDir + "Font.FarEast.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.FarEast.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("你好世界", run.GetText().Trim());
            Assert.AreEqual(2052, run.Font.LocaleIdFarEast);
            Assert.AreEqual("SimSun", run.Font.NameFarEast);
        }

        [Test]
        public void Names()
        {
            //ExStart
            //ExFor:Font.NameAscii
            //ExFor:Font.NameOther
            //ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify a font to use for all characters that fall within the ASCII character set
            builder.Font.NameAscii = "Calibri";

            // Specify a font to use for all other characters
            // This font should have a glyph for every other required character code
            builder.Font.NameOther = "Courier New";

            // The builder's font is the ASCII font
            Assert.AreEqual("Calibri", builder.Font.Name);

            // Insert a run with one word consisting of ASCII characters, and one word with all characters outside that range
            // This will create a run with two fonts
            builder.Writeln("Hello, Привет");

            doc.Save(ArtifactsDir + "Font.Names.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.Names.docx");
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
            //ExSummary:Shows how to use style name or identifier to find text formatted with a specific character style and apply different character style.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text with two styles that will be replaced by another style
            builder.Font.StyleIdentifier = StyleIdentifier.Emphasis;
            builder.Writeln("Text originally in \"Emphasis\" style");
            builder.Font.StyleIdentifier = StyleIdentifier.IntenseEmphasis;
            builder.Writeln("Text originally in \"Intense Emphasis\" style");
       
            // Loop through every run node
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true).OfType<Run>())
            {
                // If the run's text is of the "Emphasis" style, referenced by name, change the style to "Strong"
                if (run.Font.StyleName.Equals("Emphasis"))
                    run.Font.StyleName = "Strong";

                // If the run's text style is "Intense Emphasis", change it to "Strong" also, but this time reference using a StyleIdentifier
                if (run.Font.StyleIdentifier.Equals(StyleIdentifier.IntenseEmphasis))
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
        public void Style()
        {
            //ExStart
            //ExFor:Font.Style
            //ExFor:Style.BuiltIn
            //ExSummary:Applies double underline to all runs in a document that are formatted with custom character styles.
            //Document doc = new Document(MyDir + "Custom style.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a custom style
            Style style = doc.Styles.Add(StyleType.Character, "MyStyle");
            style.Font.Color = Color.Red;
            style.Font.Name = "Courier New";

            // Set the style of the current paragraph to our custom style
            // This will apply to only the text after the style separator
            builder.Font.StyleName = "MyStyle";
            builder.Write("This text is in a custom style.");
            
            // Iterate through every run node and apply underlines to the run if its style is not a built in style,
            // like the one we added
            foreach (Node node in doc.GetChildNodes(NodeType.Run, true))
            {
                Run run = (Run)node;
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
        public void SubstitutionNotification()
        {
            // Store the font sources currently used so we can restore them later
            FontSourceBase[] origFontSources = FontSettings.DefaultInstance.GetFontsSources();

            //ExStart
            //ExFor:IWarningCallback
            //ExFor:DocumentBase.WarningCallback
            //ExFor:Fonts.FontSettings.DefaultInstance
            //ExSummary:Demonstrates how to receive notifications of font substitutions by using IWarningCallback.
            // Load the document to render
            Document doc = new Document(MyDir + "Document.docx");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
            HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = callback;

            // We can choose the default font to use in the case of any missing fonts
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

            // For testing we will set Aspose.Words to look for fonts only in a folder which does not exist. Since Aspose.Words won't
            // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
            // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
            FontSettings.DefaultInstance.SetFontsFolder(string.Empty, false);

            // Pass the save options along with the save path to the save method
            doc.Save(ArtifactsDir + "Font.SubstitutionNotification.pdf");
            //ExEnd

            Assert.Greater(callback.FontWarnings.Count, 0);
            Assert.True(callback.FontWarnings[0].WarningType == WarningType.FontSubstitution);
            Assert.True(callback.FontWarnings[0].Description
                .Equals(
                    "Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));

            // Restore default fonts
            FontSettings.DefaultInstance.SetFontsSources(origFontSources);
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
            //ExSummary:Shows how to get available fonts and information about them.
            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts
            FontSourceBase[] folderFontSource = { new FolderFontSource(FontsDir, true) };
            
            foreach (PhysicalFontInfo fontInfo in folderFontSource[0].GetAvailableFonts())
            {
                Console.WriteLine("FontFamilyName : {0}", fontInfo.FontFamilyName);
                Console.WriteLine("FullFontName  : {0}", fontInfo.FullFontName);
                Console.WriteLine("Version  : {0}", fontInfo.Version);
                Console.WriteLine("FilePath : {0}\n", fontInfo.FilePath);
            }
            //ExEnd

            Assert.AreEqual(folderFontSource[0].GetAvailableFonts().Count, Directory.GetFiles(FontsDir).Count(f => f.EndsWith(".ttf")));
        }

        //ExStart
        //ExFor:Fonts.FontInfoSubstitutionRule
        //ExFor:Fonts.FontSubstitutionSettings.FontInfoSubstitution
        //ExFor:IWarningCallback
        //ExFor:IWarningCallback.Warning(WarningInfo)
        //ExFor:WarningInfo
        //ExFor:WarningInfo.Description
        //ExFor:WarningInfo.WarningType
        //ExFor:WarningInfoCollection
        //ExFor:WarningInfoCollection.Warning(WarningInfo)
        //ExFor:WarningInfoCollection.GetEnumerator
        //ExFor:WarningInfoCollection.Clear
        //ExFor:WarningType
        //ExFor:DocumentBase.WarningCallback
        //ExSummary:Shows how to set the property for finding the closest match font among the available font sources instead missing font.
        [Test]
        public void EnableFontSubstitution()
        {
            Document doc = new Document(MyDir + "Missing font.docx");

            // Assign a custom warning callback
            HandleDocumentSubstitutionWarnings substitutionWarningHandler = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = substitutionWarningHandler;

            // Set a default font name and enable font substitution
            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial"; ;
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

            // When saving the document with the missing font, we should get a warning
            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "Font.EnableFontSubstitution.pdf");

            // List all warnings using an enumerator
            using (IEnumerator<WarningInfo> warnings = substitutionWarningHandler.FontWarnings.GetEnumerator()) 
                while (warnings.MoveNext()) 
                    Console.WriteLine(warnings.Current.Description);

            // Warnings are stored in this format
            Assert.AreEqual(WarningSource.Layout, substitutionWarningHandler.FontWarnings[0].Source);
            Assert.AreEqual("Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document.", 
                substitutionWarningHandler.FontWarnings[0].Description);

            // The warning info collection can also be cleared like this
            substitutionWarningHandler.FontWarnings.Clear();

            Assert.That(substitutionWarningHandler.FontWarnings, Is.Empty);
        }

        public class HandleDocumentSubstitutionWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted
                if (info.WarningType == WarningType.FontSubstitution)
                    FontWarnings.Warning(info);
            }

            public WarningInfoCollection FontWarnings = new WarningInfoCollection();
        }
        //ExEnd

        [Test]
        public void DisableFontSubstitution()
        {
            Document doc = new Document(MyDir + "Missing font.docx");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
            HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;

            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "Font.DisableFontSubstitution.pdf");

            Regex reg = new Regex("Font '28 Days Later' has not been found. Using (.*) font instead. Reason: default font setting.");
            
            foreach (WarningInfo fontWarning in callback.FontWarnings)
            {        
                Match match = reg.Match(fontWarning.Description);
                if (match.Success)
                {
                    Assert.Pass();
                    break;
                }
            }
        }

        [Test]
        [Category("SkipMono")]
        public void SubstitutionWarnings()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
            HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SetFontsFolder(FontsDir, false);
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Arvo", "Slab");
            
            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "Font.SubstitutionWarnings.pdf");

            Assert.AreEqual("Font \'Arial\' has not been found. Using \'Arvo\' font instead. Reason: table substitution.",
                callback.FontWarnings[0].Description);
            Assert.AreEqual("Font \'Times New Roman\' has not been found. Using \'M+ 2m\' font instead. Reason: font info substitution.",
                callback.FontWarnings[1].Description);
        }

        [Test]
        public void SubstitutionWarningsClosestMatch()
        {
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
            HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = callback;

            doc.Save(ArtifactsDir + "Font.SubstitutionWarningsClosestMatch.pdf");

            Assert.True(callback.FontWarnings[0].Description
                .Equals("Font \'SymbolPS\' has not been found. Using \'Wingdings\' font instead. Reason: font info substitution."));
        }

        [Test]
        public void SetFontAutoColor()
        {
            //ExStart
            //ExFor:Font.AutoColor
            //ExSummary:Shows how calculated color of the text (black or white) to be used for 'auto color'
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Remove direct color, so it can be calculated automatically with Font.AutoColor
            builder.Font.Color = Color.Empty;

            // When we set black color for background, autocolor for font must be white
            builder.Font.Shading.BackgroundPatternColor = Color.Black;

            builder.Writeln("The text color automatically chosen for this run is white.");

            // When we set white color for background, autocolor for font must be black
            builder.Font.Shading.BackgroundPatternColor = Color.White;

            builder.Writeln("The text color automatically chosen for this run is black.");

            doc.Save(ArtifactsDir + "Font.SetFontAutoColor.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Font.SetFontAutoColor.docx");
            Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("The text color automatically chosen for this run is white.", run.GetText().Trim());
            Assert.AreEqual(Color.Empty.ToArgb(), run.Font.Color.ToArgb());
            Assert.AreEqual(Color.Black.ToArgb(), run.Font.Shading.BackgroundPatternColor.ToArgb());

            run = doc.FirstSection.Body.Paragraphs[1].Runs[0];

            Assert.AreEqual("The text color automatically chosen for this run is black.", run.GetText().Trim());
            Assert.AreEqual(Color.Empty.ToArgb(), run.Font.Color.ToArgb());
            Assert.AreEqual(Color.White.ToArgb(), run.Font.Shading.BackgroundPatternColor.ToArgb());
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
        //ExSummary:Implements the Visitor Pattern to remove all content formatted as hidden from the document.
        [Test] //ExSkip
        public void RemoveHiddenContentFromDocument()
        {
            // Open the document we want to remove hidden content from
            Document doc = new Document(MyDir + "Hidden content.docx");
            Assert.AreEqual(26, doc.GetChildNodes(NodeType.Paragraph, true).Count); //ExSkip
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count); //ExSkip

            // Create an object that inherits from the DocumentVisitor class
            RemoveHiddenContentVisitor hiddenContentRemover = new RemoveHiddenContentVisitor();

            // We can run it over the entire the document like so
            doc.Accept(hiddenContentRemover);

            // Or we can run it on only a specific node
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 4, true);
            para.Accept(hiddenContentRemover);

            // Or over a different type of node like below
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            table.Accept(hiddenContentRemover);

            doc.Save(ArtifactsDir + "Font.RemoveHiddenContentFromDocument.docx");
            TestRemoveHiddenContent(new Document(ArtifactsDir + "Font.RemoveHiddenContentFromDocument.docx")); //ExSkip
        }

        /// <summary>
        /// This class when executed will remove all hidden content from the Document. Implemented as a Visitor.
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
                // Currently there is no way to tell if a particular Table/Row/Cell is hidden. 
                // Instead, if the content of a table is hidden, then all inline child nodes of the table should be 
                // hidden and thus removed by previous visits as well. This will result in the container being empty
                // If this is the case, we know to remove the table node.
                //
                // Note that a table which is not hidden but simply has no content will not be affected by this algorithm,
                // as technically they are not completely empty (for example a properly formed Cell will have at least 
                // an empty paragraph in it)
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
        public void BlankDocumentFonts()
        {
            //ExStart
            //ExFor:Fonts.FontInfoCollection.Contains(String)
            //ExFor:Fonts.FontInfoCollection.Count
            //ExSummary:Shows info about the fonts that are present in the blank document.
            Document doc = new Document();

            // A blank document comes with 3 fonts
            Assert.AreEqual(3, doc.FontInfos.Count);

            // Their names can be looked up like this
            Assert.AreEqual("Times New Roman", doc.FontInfos[0].Name);
            Assert.AreEqual("Symbol", doc.FontInfos[1].Name);
            Assert.AreEqual("Arial", doc.FontInfos[2].Name);
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
            //ExSummary:Shows how to extract embedded font from a document.
            Document doc = new Document(MyDir + "Embedded font.docx");

            // Get the FontInfo for the embedded font
            FontInfo embeddedFont = doc.FontInfos["Alte DIN 1451 Mittelschrift"];

            // We can now extract this embedded font
            byte[] embeddedFontBytes = embeddedFont.GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular);
            Assert.IsNotNull(embeddedFontBytes);

            // Then we can save the font to our directory
            File.WriteAllBytes(ArtifactsDir + "Alte DIN 1451 Mittelschrift.ttf", embeddedFontBytes);
            
            // If we want to extract a font from a .doc as opposed to a .docx, we need to make sure to set the appropriate embedded font format
            doc = new Document(MyDir + "Embedded font.doc");

            Assert.IsNull(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular));
            Assert.IsNotNull(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFont(EmbeddedFontFormat.EmbeddedOpenType, EmbeddedFontStyle.Regular));
            // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType
            Assert.IsNotNull(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFontAsOpenType(EmbeddedFontStyle.Regular));
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
            //ExSummary:Shows how to get information about each font in a document.
            Document doc = new Document(MyDir + "Document.docx");
            
            // We can iterate over all the fonts with an enumerator
            IEnumerator fontCollectionEnumerator = doc.FontInfos.GetEnumerator();
            // Print detailed information about each font to the console
            while (fontCollectionEnumerator.MoveNext())
            {
                FontInfo fontInfo = (FontInfo)fontCollectionEnumerator.Current;
                if (fontInfo != null)
                {
                    Console.WriteLine("Font name: " + fontInfo.Name);
                    // Alt names are usually blank
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

            Assert.AreEqual(new int[] { 2, 15, 5, 2, 2, 2, 4, 3, 2, 4 }, doc.FontInfos["Calibri"].Panose);
            Assert.AreEqual(new int[] { 2, 2, 6, 3, 5, 4, 5, 2, 3, 4 }, doc.FontInfos["Times New Roman"].Panose);
        }

        [Test]
        public void FontSourceFile()
        {
            //ExStart
            //ExFor:Fonts.FileFontSource
            //ExFor:Fonts.FileFontSource.#ctor(String)
            //ExFor:Fonts.FileFontSource.#ctor(String, Int32)
            //ExFor:Fonts.FileFontSource.FilePath
            //ExFor:Fonts.FileFontSource.Type
            //ExFor:Fonts.FontSourceBase
            //ExFor:Fonts.FontSourceBase.Priority
            //ExFor:Fonts.FontSourceBase.Type
            //ExFor:Fonts.FontSourceType
            //ExSummary:Shows how to create a file font source.
            Document doc = new Document();

            // Create a font settings object for our document
            doc.FontSettings = new FontSettings();

            // Create a font source from a file in our system
            FileFontSource fileFontSource = new FileFontSource(MyDir + "Alte DIN 1451 Mittelschrift.ttf", 0);

            // Import the font source into our document
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { fileFontSource });

            Assert.AreEqual(MyDir + "Alte DIN 1451 Mittelschrift.ttf", fileFontSource.FilePath);
            Assert.AreEqual(FontSourceType.FontFile, fileFontSource.Type);
            Assert.AreEqual(0, fileFontSource.Priority);
            //ExEnd
        }

        [Test]
        public void FontSourceFolder()
        {
            //ExStart
            //ExFor:Fonts.FolderFontSource
            //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean)
            //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean, Int32)
            //ExFor:Fonts.FolderFontSource.FolderPath
            //ExFor:Fonts.FolderFontSource.ScanSubfolders
            //ExFor:Fonts.FolderFontSource.Type
            //ExSummary:Shows how to create a folder font source.
            Document doc = new Document();

            // Create a font settings object for our document
            doc.FontSettings = new FontSettings();

            // Create a font source from a folder that contains font files
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false, 1);

            // Add that source to our document
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { folderFontSource });

            Assert.AreEqual(FontsDir, folderFontSource.FolderPath);
            Assert.AreEqual(false, folderFontSource.ScanSubfolders);
            Assert.AreEqual(FontSourceType.FontsFolder, folderFontSource.Type);
            Assert.AreEqual(1, folderFontSource.Priority);
            //ExEnd
        }

        [Test]
        public void FontSourceMemory()
        {
            //ExStart
            //ExFor:Fonts.MemoryFontSource
            //ExFor:Fonts.MemoryFontSource.#ctor(Byte[])
            //ExFor:Fonts.MemoryFontSource.#ctor(Byte[], Int32)
            //ExFor:Fonts.MemoryFontSource.FontData
            //ExFor:Fonts.MemoryFontSource.Type
            //ExSummary:Shows how to create a memory font source.
            Document doc = new Document();

            // Create a font settings object for our document
            doc.FontSettings = new FontSettings();

            // Import a font file, putting its contents into a byte array
            byte[] fontBytes = File.ReadAllBytes(MyDir + "Alte DIN 1451 Mittelschrift.ttf");

            // Create a memory font source from our array
            MemoryFontSource memoryFontSource = new MemoryFontSource(fontBytes, 0);

            // Add that font source to our document
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { memoryFontSource });

            Assert.AreEqual(FontSourceType.MemoryFont, memoryFontSource.Type);
            Assert.AreEqual(0, memoryFontSource.Priority);
            //ExEnd
        }

        [Test]
        public void FontSourceSystem()
        {
            //ExStart
            //ExFor:TableSubstitutionRule.AddSubstitutes(String, String[])
            //ExFor:FontSubstitutionRule.Enabled
            //ExFor:TableSubstitutionRule.GetSubstitutes(String)
            //ExFor:Fonts.FontSettings.ResetFontSources
            //ExFor:Fonts.FontSettings.SubstitutionSettings
            //ExFor:Fonts.FontSubstitutionSettings
            //ExFor:Fonts.SystemFontSource
            //ExFor:Fonts.SystemFontSource.#ctor
            //ExFor:Fonts.SystemFontSource.#ctor(Int32)
            //ExFor:Fonts.SystemFontSource.GetSystemFontFolders
            //ExFor:Fonts.SystemFontSource.Type
            //ExSummary:Shows how to access a document's system font source and set font substitutes.
            Document doc = new Document();

            // Create a font settings object for our document
            doc.FontSettings = new FontSettings();

            // By default, we always start with a system font source
            Assert.AreEqual(1, doc.FontSettings.GetFontsSources().Length);

            SystemFontSource systemFontSource = (SystemFontSource)doc.FontSettings.GetFontsSources()[0];
            Assert.AreEqual(FontSourceType.SystemFonts, systemFontSource.Type);
            Assert.AreEqual(0, systemFontSource.Priority);
            
            PlatformID pid = Environment.OSVersion.Platform;
            bool isWindows = (pid == PlatformID.Win32NT) || (pid == PlatformID.Win32S) || (pid == PlatformID.Win32Windows) || (pid == PlatformID.WinCE);
            if (isWindows)
            {
                const string fontsPath = @"C:\WINDOWS\Fonts";
                Assert.AreEqual(fontsPath.ToLower(), SystemFontSource.GetSystemFontFolders().FirstOrDefault()?.ToLower());
            }

            foreach (string systemFontFolder in SystemFontSource.GetSystemFontFolders())
            {
                Console.WriteLine(systemFontFolder);
            }

            // Set a font that exists in the Windows Fonts directory as a substitute for one that doesn't
            doc.FontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;
            doc.FontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Kreon-Regular", new string[] { "Calibri" });

            Assert.AreEqual(1, doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").Count());
            Assert.Contains("Calibri", doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").ToArray());

            // Alternatively, we could add a folder font source in which the corresponding folder contains the font
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false);
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { systemFontSource, folderFontSource });
            Assert.AreEqual(2, doc.FontSettings.GetFontsSources().Length);

            // Resetting the font sources still leaves us with the system font source as well as our substitutes
            doc.FontSettings.ResetFontSources();

            Assert.AreEqual(1, doc.FontSettings.GetFontsSources().Length);
            Assert.AreEqual(FontSourceType.SystemFonts, doc.FontSettings.GetFontsSources()[0].Type);
            Assert.AreEqual(1, doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").Count());
            //ExEnd
        }

        [Test]
        public void LoadFontFallbackSettingsFromFile()
        {
            //ExStart
            //ExFor:FontFallbackSettings.Load(String)
            //ExFor:FontFallbackSettings.Save(String)
            //ExSummary:Shows how to load and save font fallback settings from file.
            Document doc = new Document(MyDir + "Rendering.docx");
            
            // By default, fallback settings are initialized with predefined settings which mimics the Microsoft Word fallback
            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.Load(MyDir + "Font fallback rules.xml");

            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "Font.LoadFontFallbackSettingsFromFile.pdf");

            // Saves font fallback setting by string
            doc.FontSettings.FallbackSettings.Save(ArtifactsDir + "FallbackSettings.xml");
            //ExEnd
        }

        [Test]
        public void LoadFontFallbackSettingsFromStream()
        {
            //ExStart
            //ExFor:FontFallbackSettings.Load(Stream)
            //ExFor:FontFallbackSettings.Save(Stream)
            //ExSummary:Shows how to load and save font fallback settings from stream.
            Document doc = new Document(MyDir + "Rendering.docx");

            // By default, fallback settings are initialized with predefined settings which mimics the Microsoft Word fallback
            using (FileStream fontFallbackStream = new FileStream(MyDir + "Font fallback rules.xml", FileMode.Open))
            {
                FontSettings fontSettings = new FontSettings();
                fontSettings.FallbackSettings.Load(fontFallbackStream);

                doc.FontSettings = fontSettings;
            }

            doc.Save(ArtifactsDir + "Font.LoadFontFallbackSettingsFromStream.pdf");

            // Saves font fallback setting by stream
            using (FileStream fontFallbackStream =
                new FileStream(ArtifactsDir + "FallbackSettings.xml", FileMode.Create))
            {
                doc.FontSettings.FallbackSettings.Save(fontFallbackStream);
            }
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "FallbackSettings.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules = fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.AreEqual("0B80-0BFF", rules[0].Attributes["Ranges"].Value);
            Assert.AreEqual("Vijaya", rules[0].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("1F300-1F64F", rules[1].Attributes["Ranges"].Value);
            Assert.AreEqual("Segoe UI Emoji, Segoe UI Symbol", rules[1].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("2000-206F, 2070-209F, 20B9", rules[2].Attributes["Ranges"].Value);
            Assert.AreEqual("Arial", rules[2].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("3040-309F", rules[3].Attributes["Ranges"].Value);
            Assert.AreEqual("MS Gothic", rules[3].Attributes["FallbackFonts"].Value);
            Assert.AreEqual("Times New Roman", rules[3].Attributes["BaseFonts"].Value);

            Assert.AreEqual("3040-309F", rules[4].Attributes["Ranges"].Value);
            Assert.AreEqual("MS Mincho", rules[4].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("Arial Unicode MS", rules[5].Attributes["FallbackFonts"].Value);
        }

        [Test]
        public void LoadNotoFontsFallbackSettings()
        {
            //ExStart
            //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
            //ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
            FontSettings fontSettings = new FontSettings();

            // These are free fonts licensed under SIL OFL
            // They can be downloaded from https://www.google.com/get/noto/#sans-lgc
            fontSettings.SetFontsFolder(FontsDir + "Noto", false);

            // Note that only Sans style Noto fonts with regular weight are used in the predefined settings
            // Some of the Noto fonts uses advanced typography features
            // Advanced typography is currently not supported by AW and these fonts may be rendered inaccurately
            fontSettings.FallbackSettings.LoadNotoFallbackSettings();
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Noto Sans";

            Document doc = new Document();
            doc.FontSettings = fontSettings;
            //ExEnd
        }

        [Test]
        public void DefaultFontSubstitutionRule()
        {
            //ExStart
            //ExFor:Fonts.DefaultFontSubstitutionRule
            //ExFor:Fonts.DefaultFontSubstitutionRule.DefaultFontName
            //ExFor:Fonts.FontSubstitutionSettings.DefaultFontSubstitution
            //ExSummary:Shows how to set the default font substitution rule.
            // Create a blank document and a new FontSettings property
            Document doc = new Document();
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Get the default substitution rule within FontSettings, which will be enabled by default and will substitute all missing fonts with "Times New Roman"
            DefaultFontSubstitutionRule defaultFontSubstitutionRule = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
            Assert.True(defaultFontSubstitutionRule.Enabled);
            Assert.AreEqual("Times New Roman", defaultFontSubstitutionRule.DefaultFontName);

            // Set the default font substitute to "Courier New"
            defaultFontSubstitutionRule.DefaultFontName = "Courier New";

            // Using a document builder, add some text in a font that we do not have to see the substitution take place,
            // and render the result in a PDF
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Missing Font";
            builder.Writeln("Line written in a missing font, which will be substituted with Courier New.");

            doc.Save(ArtifactsDir + "Font.DefaultFontSubstitutionRule.pdf");
            //ExEnd

            Assert.AreEqual("Courier New", defaultFontSubstitutionRule.DefaultFontName);
        }

        [Test]
        public void FontConfigSubstitution()
        {
            //ExStart
            //ExFor:Fonts.FontConfigSubstitutionRule
            //ExFor:Fonts.FontConfigSubstitutionRule.Enabled
            //ExFor:Fonts.FontConfigSubstitutionRule.IsFontConfigAvailable
            //ExFor:Fonts.FontConfigSubstitutionRule.ResetCache
            //ExFor:Fonts.FontSubstitutionRule
            //ExFor:Fonts.FontSubstitutionRule.Enabled
            //ExFor:Fonts.FontSubstitutionSettings.FontConfigSubstitution
            //ExSummary:Shows OS-dependent font config substitution.
            // Create a new FontSettings object and get its font config substitution rule
            FontSettings fontSettings = new FontSettings();
            FontConfigSubstitutionRule fontConfigSubstitution = fontSettings.SubstitutionSettings.FontConfigSubstitution;

            bool isWindows = new[] { PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE }
                .Any(p => Environment.OSVersion.Platform == p);

            // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms
            // On Windows, it is unavailable
            if (isWindows)
            {
                Assert.False(fontConfigSubstitution.Enabled);
                Assert.False(fontConfigSubstitution.IsFontConfigAvailable());
            }

            bool isLinuxOrMac = new[] { PlatformID.Unix, PlatformID.MacOSX }.Any(p => Environment.OSVersion.Platform == p);

            // On Linux/Mac, we will have access and will be able to perform operations
            if (isLinuxOrMac)
            {
                Assert.True(fontConfigSubstitution.Enabled);
                Assert.True(fontConfigSubstitution.IsFontConfigAvailable());

                fontConfigSubstitution.ResetCache();
            }
            //ExEnd
        }

        [Test]
        public void FallbackSettings()
        {
            //ExStart
            //ExFor:Fonts.FontFallbackSettings.LoadMsOfficeFallbackSettings
            //ExFor:Fonts.FontFallbackSettings.LoadNotoFallbackSettings
            //ExSummary:Shows how to load pre-defined fallback font settings.
            Document doc = new Document();

            // Create a FontSettings object for our document and get its FallbackSettings attribute
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;
            FontFallbackSettings fontFallbackSettings = fontSettings.FallbackSettings;

            // Save the default fallback font scheme in an XML document
            // For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts
            // This means that if the font we are using does not have symbols for the 0x0C00-0x0C7F Unicode block,
            // the symbols from the "Vani" font will be used as a substitute
            fontFallbackSettings.Save(ArtifactsDir + "Font.FallbackSettings.Default.xml");

            // There are two pre-defined font fallback schemes we can choose from
            // 1: Use the default Microsoft Office scheme, which is the same one as the default
            fontFallbackSettings.LoadMsOfficeFallbackSettings();
            fontFallbackSettings.Save(ArtifactsDir + "Font.FallbackSettings.LoadMsOfficeFallbackSettings.xml");

            // 2: Use the scheme built from Google Noto fonts
            fontFallbackSettings.LoadNotoFallbackSettings();
            fontFallbackSettings.Save(ArtifactsDir + "Font.FallbackSettings.LoadNotoFallbackSettings.xml");
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "Font.FallbackSettings.Default.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules = fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.AreEqual("0C00-0C7F", rules[5].Attributes["Ranges"].Value);
            Assert.AreEqual("Vani", rules[5].Attributes["FallbackFonts"].Value);
        }

        [Test]
        public void FallbackSettingsCustom()
        {
            //ExStart
            //ExFor:Fonts.FontSettings.FallbackSettings
            //ExFor:Fonts.FontFallbackSettings
            //ExFor:Fonts.FontFallbackSettings.BuildAutomatic
            //ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
            Document doc = new Document();

            // Create a FontSettings object for our document and get its FallbackSettings attribute
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;
            FontFallbackSettings fontFallbackSettings = fontSettings.FallbackSettings;

            // Set our fonts to be sourced exclusively from the "MyFonts" folder
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false);
            fontSettings.SetFontsSources(new FontSourceBase[] { folderFontSource });

            // Calling BuildAutomatic() will generate a fallback scheme that distributes accessible fonts across as many Unicode character codes as possible
            // In our case, it only has access to the handful of fonts inside the "MyFonts" folder
            fontFallbackSettings.BuildAutomatic();
            fontFallbackSettings.Save(ArtifactsDir + "Font.FallbackSettingsCustom.BuildAutomatic.xml");

            // We can also load a custom substitution scheme from a file like this
            // This scheme applies the "Arvo" font across the "0000-00ff" Unicode blocks, the "Squarish Sans CT" font across "0100-024f",
            // and the "M+ 2m" font in every place that none of the other fonts cover
            fontFallbackSettings.Load(MyDir + "Custom font fallback settings.xml");

            // Create a document builder and set its font to one that does not exist in any of our sources
            // In doing that we will rely completely on our font fallback scheme to render text
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Missing Font";

            // Type out every Unicode character from 0x0021 to 0x052F, with descriptive lines dividing Unicode blocks we defined in our custom font fallback scheme
            for (int i = 0x0021; i < 0x0530; i++)
            {
                switch (i)
                {
                    case 0x0021:
                        builder.Writeln("\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement Unicode blocks in \"Arvo\" font:");
                        break;
                    case 0x0100:
                        builder.Writeln("\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in \"Squarish Sans CT\" font:");
                        break;
                    case 0x0250:
                        builder.Writeln("\n\n0x0250 - 0x052F: \nIPA/Greek/Cyrillic blocks in \"M+ 2m\" font:");
                        break;
                }

                builder.Write($"{Convert.ToChar(i)}");
            }

            doc.Save(ArtifactsDir + "Font.FallbackSettingsCustom.pdf");
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "Font.FallbackSettingsCustom.BuildAutomatic.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules = fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.AreEqual("0000-007F", rules[0].Attributes["Ranges"].Value);
            Assert.AreEqual("Arvo", rules[0].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("0180-024F", rules[3].Attributes["Ranges"].Value);
            Assert.AreEqual("M+ 2m", rules[3].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("0300-036F", rules[6].Attributes["Ranges"].Value);
            Assert.AreEqual("Noticia Text", rules[6].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("0590-05FF", rules[10].Attributes["Ranges"].Value);
            Assert.AreEqual("Squarish Sans CT", rules[10].Attributes["FallbackFonts"].Value);
        }

        [Test]
        public void TableSubstitutionRule()
        {
            //ExStart
            //ExFor:Fonts.TableSubstitutionRule
            //ExFor:Fonts.TableSubstitutionRule.LoadLinuxSettings
            //ExFor:Fonts.TableSubstitutionRule.LoadWindowsSettings
            //ExFor:Fonts.TableSubstitutionRule.Save(System.IO.Stream)
            //ExFor:Fonts.TableSubstitutionRule.Save(System.String)
            //ExSummary:Shows how to access font substitution tables for Windows and Linux.
            // Create a blank document and a new FontSettings object
            Document doc = new Document();
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Create a new table substitution rule and load the default Windows font substitution table
            TableSubstitutionRule tableSubstitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;
            tableSubstitutionRule.LoadWindowsSettings();

            // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman"
            Assert.AreEqual(new[] { "Times New Roman" }, tableSubstitutionRule.GetSubstitutes("Times New Roman CE").ToArray());

            // We can save the table for viewing in the form of an XML document
            tableSubstitutionRule.Save(ArtifactsDir + "Font.TableSubstitutionRule.Windows.xml");

            // Linux has its own substitution table
            // If "FreeSerif" is unavailable to substitute for "Times New Roman CE", we then look for "Liberation Serif", and so on
            tableSubstitutionRule.LoadLinuxSettings();
            Assert.AreEqual(new[] { "FreeSerif", "Liberation Serif", "DejaVu Serif" }, tableSubstitutionRule.GetSubstitutes("Times New Roman CE").ToArray());

            // Save the Linux substitution table using a stream
            using (FileStream fileStream = new FileStream(ArtifactsDir + "Font.TableSubstitutionRule.Linux.xml", FileMode.Create))
            {
                tableSubstitutionRule.Save(fileStream);
            }
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "Font.TableSubstitutionRule.Windows.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules = fallbackSettingsDoc.SelectNodes("//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

            Assert.AreEqual("Times New Roman CE", rules[16].Attributes["OriginalFont"].Value);
            Assert.AreEqual("Times New Roman", rules[16].Attributes["SubstituteFonts"].Value);

            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "Font.TableSubstitutionRule.Linux.xml"));
            rules = fallbackSettingsDoc.SelectNodes("//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

            Assert.AreEqual("Times New Roman CE", rules[31].Attributes["OriginalFont"].Value);
            Assert.AreEqual("FreeSerif, Liberation Serif, DejaVu Serif", rules[31].Attributes["SubstituteFonts"].Value);
        }

        [Test]
        public void TableSubstitutionRuleCustom()
        {
            //ExStart
            //ExFor:Fonts.FontSubstitutionSettings.TableSubstitution
            //ExFor:Fonts.TableSubstitutionRule.AddSubstitutes(System.String,System.String[])
            //ExFor:Fonts.TableSubstitutionRule.GetSubstitutes(System.String)
            //ExFor:Fonts.TableSubstitutionRule.Load(System.IO.Stream)
            //ExFor:Fonts.TableSubstitutionRule.Load(System.String)
            //ExFor:Fonts.TableSubstitutionRule.SetSubstitutes(System.String,System.String[])
            //ExSummary:Shows how to work with custom font substitution tables.
            // Create a blank document and a new FontSettings object
            Document doc = new Document();
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Create a new table substitution rule and load the default Windows font substitution table
            TableSubstitutionRule tableSubstitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;

            // If we select fonts exclusively from our own folder, we will need a custom substitution table
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false);
            fontSettings.SetFontsSources(new FontSourceBase[] { folderFontSource });

            // There are two ways of loading a substitution table from a file in the local file system
            // 1: Loading from a stream
            using (FileStream fileStream = new FileStream(MyDir + "Font substitution rules.xml", FileMode.Open))
            {
                tableSubstitutionRule.Load(fileStream);
            }

            // 2: Load directly from file
            tableSubstitutionRule.Load(MyDir + "Font substitution rules.xml");

            // Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font", which we don't have,
            // and then with "Kreon", found in the "MyFonts" folder
            Assert.AreEqual(new[] { "Missing Font", "Kreon" }, tableSubstitutionRule.GetSubstitutes("Arial").ToArray());

            // If we find this substitution table lacking, we can also expand it programmatically
            // In this case, we add an entry that substitutes "Times New Roman" with "Arvo"
            Assert.Null(tableSubstitutionRule.GetSubstitutes("Times New Roman"));
            tableSubstitutionRule.AddSubstitutes("Times New Roman", "Arvo");
            Assert.AreEqual(new[] { "Arvo" }, tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray());

            // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes()
            // In case "Arvo" is unavailable, our table will look for "M+ 2m"
            tableSubstitutionRule.AddSubstitutes("Times New Roman", "M+ 2m");
            Assert.AreEqual(new[] { "Arvo", "M+ 2m" }, tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray());

            // SetSubstitutes() can set a new list of substitute fonts for a font
            tableSubstitutionRule.SetSubstitutes("Times New Roman", new[] { "Squarish Sans CT", "M+ 2m" });
            Assert.AreEqual(new[] { "Squarish Sans CT", "M+ 2m" }, tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray());

            // TO demonstrate substitution, write text in fonts we have no access to and render the result in a PDF
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Arial";
            builder.Writeln("Text written in Arial, to be substituted by Kreon.");

            builder.Font.Name = "Times New Roman";
            builder.Writeln("Text written in Times New Roman, to be substituted by Squarish Sans CT.");

            doc.Save(ArtifactsDir + "Font.TableSubstitutionRule.Custom.pdf");
            //ExEnd
        }

        [Test]
        public void ResolveFontsBeforeLoadingDocument()
        {
            //ExStart
            //ExFor:LoadOptions.FontSettings
            //ExSummary:Shows how to designate font substitutes during loading.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = new FontSettings();

            // Set a font substitution rule for a LoadOptions object that replaces a font that's not installed in our system with one that is
            TableSubstitutionRule substitutionRule = loadOptions.FontSettings.SubstitutionSettings.TableSubstitution;
            substitutionRule.AddSubstitutes("MissingFont", new string[] { "Comic Sans MS" });

            // If we pass that object while loading a document, any text with the "MissingFont" font will change to "Comic Sans MS"
            Document doc = new Document(MyDir + "Missing font.html", loadOptions);

            // At this point such text will still be in "MissingFont", and font substitution will be carried out once we save
            Assert.AreEqual("MissingFont", doc.FirstSection.Body.FirstParagraph.Runs[0].Font.Name);

            doc.Save(ArtifactsDir + "Font.ResolveFontsBeforeLoadingDocument.pdf");
            //ExEnd
        }
        
        [Test]
        public void LineSpacing()
        {
            //ExStart
            //ExFor:Font.LineSpacing
            //ExSummary:Shows how to get line spacing of current font (in points).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set different fonts for the DocumentBuilder and verify their line spacing
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
            //ExSummary:Shows how to checks if particular Dml text effect is applied.
            Document doc = new Document(MyDir + "DrawingML text effects.docx");
            
            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;
            
            Assert.True(runs[0].Font.HasDmlEffect(TextDmlEffect.Shadow));
            Assert.True(runs[1].Font.HasDmlEffect(TextDmlEffect.Shadow));
            Assert.True(runs[2].Font.HasDmlEffect(TextDmlEffect.Reflection));
            Assert.True(runs[3].Font.HasDmlEffect(TextDmlEffect.Effect3D));
            Assert.True(runs[4].Font.HasDmlEffect(TextDmlEffect.Fill));
            //ExEnd
        }

        //ExStart
        //ExFor:StreamFontSource
        //ExFor:StreamFontSource.OpenFontDataStream
        //ExSummary:Shows how to allows to load fonts from stream.
        [Test] //ExSkip
        public void StreamFontSourceFileRendering()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsSources(new FontSourceBase[] { new StreamFontSourceFile() });

            DocumentBuilder builder = new DocumentBuilder();
            builder.Document.FontSettings = fontSettings;
            builder.Font.Name = "Kreon-Regular";
            builder.Writeln("Test aspose text when saving to PDF.");

            builder.Document.Save(ArtifactsDir + "Font.StreamFontSourceFileRendering.pdf");
        }
        
        /// <summary>
        /// Load the font data only when it is required and not to store it in the memory for the "FontSettings" lifetime.
        /// </summary>
        private class StreamFontSourceFile : StreamFontSource
        {
            public override Stream OpenFontDataStream()
            {
                return File.OpenRead(FontsDir + "Kreon-Regular.ttf");
            }
        }
        //ExEnd

        [Test, Category("IgnoreOnJenkins")]
        public void CheckScanUserFontsFolder()
        {
            // On Windows 10 fonts may be installed either into system folder "%windir%\fonts" for all users
            // or into user folder "%userprofile%\AppData\Local\Microsoft\Windows\Fonts" for current user
            SystemFontSource systemFontSource = new SystemFontSource();
            Assert.NotNull(systemFontSource.GetAvailableFonts()
                    .FirstOrDefault(x => x.FilePath.Contains("\\AppData\\Local\\Microsoft\\Windows\\Fonts")),
                "Fonts did not install to the user font folder");
        }
    }
}
#endif