// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExParagraphFormat : ApiExampleBase
    {
        [Test]
        public void AsianTypographyProperties()
        {
            //ExStart
            //ExFor:ParagraphFormat.FarEastLineBreakControl
            //ExFor:ParagraphFormat.WordWrap
            //ExFor:ParagraphFormat.HangingPunctuation
            //ExSummary:Shows how to set special properties for Asian typography. 
            Document doc = new Document(MyDir + "Document.docx");

            ParagraphFormat format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;
            format.FarEastLineBreakControl = true;
            format.WordWrap = false;
            format.HangingPunctuation = true;

            doc.Save(ArtifactsDir + "ParagraphFormat.AsianTypographyProperties.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.AsianTypographyProperties.docx");
            format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;

            Assert.True(format.FarEastLineBreakControl);
            Assert.False(format.WordWrap);
            Assert.True(format.HangingPunctuation);
        }

        [Test]
        public void DropCap()
        {
            //ExStart
            //ExFor:DropCapPosition
            //ExSummary:Shows how to set the position of a drop cap.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Every paragraph has its own drop cap setting
            ParagraphFormat format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;

            // By default, it is "none", for no drop caps
            Assert.AreEqual(DropCapPosition.None, format.DropCapPosition);

            // Move the first capital to outside the text margin
            format.DropCapPosition = DropCapPosition.Margin;
            format.LinesToDrop = 2;

            // This text will be affected
            builder.Write("Hello world!");

            doc.Save(ArtifactsDir + "ParagraphFormat.DropCap.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.DropCap.docx");
            format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;

            Assert.AreEqual(DropCapPosition.Margin, format.DropCapPosition);
            Assert.AreEqual(2, format.LinesToDrop);
        }

        [Test]
        public void LineSpacing()
        {
            //ExStart
            //ExFor:ParagraphFormat.LineSpacing
            //ExFor:ParagraphFormat.LineSpacingRule
            //ExSummary:Shows how to work with line spacing.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the paragraph's line spacing to have a minimum value
            // This will give vertical padding to lines of text of any size that's too small to maintain the line height
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            builder.ParagraphFormat.LineSpacing = 20.0;

            builder.Writeln("Minimum line spacing of 20.");
            builder.Writeln("Minimum line spacing of 20.");

            // Set the line spacing to always be exactly 5 points
            // If the font size is larger than the spacing, the top of the text will be truncated
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            builder.ParagraphFormat.LineSpacing = 5.0;

            builder.Writeln("Line spacing of exactly 5.");
            builder.Writeln("Line spacing of exactly 5.");

            // Set the line spacing to a multiple of the default line spacing, which is 12 points by default
            // 18 points will set the spacing to always be 1.5 lines, which will scale with different font sizes
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            builder.ParagraphFormat.LineSpacing = 18.0;

            builder.Writeln("Line spacing of 1.5 default lines.");
            builder.Writeln("Line spacing of 1.5 default lines.");

            doc.Save(ArtifactsDir + "ParagraphFormat.LineSpacing.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.LineSpacing.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(LineSpacingRule.AtLeast, paragraphs[0].ParagraphFormat.LineSpacingRule);
            Assert.AreEqual(20.0d, paragraphs[0].ParagraphFormat.LineSpacing);
            Assert.AreEqual(LineSpacingRule.AtLeast, paragraphs[1].ParagraphFormat.LineSpacingRule);
            Assert.AreEqual(20.0d, paragraphs[1].ParagraphFormat.LineSpacing);

            Assert.AreEqual(LineSpacingRule.Exactly, paragraphs[2].ParagraphFormat.LineSpacingRule);
            Assert.AreEqual(5.0d, paragraphs[2].ParagraphFormat.LineSpacing);
            Assert.AreEqual(LineSpacingRule.Exactly, paragraphs[3].ParagraphFormat.LineSpacingRule);
            Assert.AreEqual(5.0d, paragraphs[3].ParagraphFormat.LineSpacing);

            Assert.AreEqual(LineSpacingRule.Multiple, paragraphs[4].ParagraphFormat.LineSpacingRule);
            Assert.AreEqual(18.0d, paragraphs[4].ParagraphFormat.LineSpacing);
            Assert.AreEqual(LineSpacingRule.Multiple, paragraphs[5].ParagraphFormat.LineSpacingRule);
            Assert.AreEqual(18.0d, paragraphs[5].ParagraphFormat.LineSpacing);
        }

        [Test]
        public void ParagraphSpacing()
        {
            //ExStart
            //ExFor:ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle
            //ExFor:ParagraphFormat.SpaceAfter
            //ExFor:ParagraphFormat.SpaceAfterAuto
            //ExFor:ParagraphFormat.SpaceBefore
            //ExFor:ParagraphFormat.SpaceBeforeAuto
            //ExSummary:Shows how to work with paragraph spacing.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the amount of white space before and after each paragraph to 12 points
            builder.ParagraphFormat.SpaceBefore = 12.0f;
            builder.ParagraphFormat.SpaceAfter = 12.0f;

            // We can set these flags to apply default spacing, effectively ignoring the spacing in the attributes we set above
            Assert.False(builder.ParagraphFormat.SpaceAfterAuto);
            Assert.False(builder.ParagraphFormat.SpaceBeforeAuto);
            Assert.False(builder.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle);

            // Insert two paragraphs which will have padding above and below them and save the document
            builder.Writeln("Paragraph 1.");
            builder.Writeln("Paragraph 2.");

            doc.Save(ArtifactsDir + "ParagraphFormat.ParagraphSpacing.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.ParagraphSpacing.docx");
            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;

            Assert.AreEqual(12.0d, format.SpaceBefore);
            Assert.AreEqual(12.0d, format.SpaceAfter);
            Assert.False(format.SpaceAfterAuto);
            Assert.False(format.SpaceBeforeAuto);
            Assert.False(format.NoSpaceBetweenParagraphsOfSameStyle);

            format = doc.FirstSection.Body.Paragraphs[1].ParagraphFormat;

            Assert.AreEqual(12.0d, format.SpaceBefore);
            Assert.AreEqual(12.0d, format.SpaceAfter);
            Assert.False(format.SpaceAfterAuto);
            Assert.False(format.SpaceBeforeAuto);
            Assert.False(format.NoSpaceBetweenParagraphsOfSameStyle);
        }

        [Test]
        public void ParagraphOutlineLevel()
        {
            //ExStart
            //ExFor:ParagraphFormat.OutlineLevel
            //ExSummary:Shows how to set paragraph outline levels to create collapsible text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Each paragraph has an OutlineLevel, which could be any number from 1 to 9, or at the default "BodyText" value
            // Setting the attribute to one of the numbered values will enable an arrow in Microsoft Word
            // next to the beginning of the paragraph that, when clicked, will collapse the paragraph
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
            builder.Writeln("Paragraph outline level 1.");

            // Level 1 is the topmost level, which practically means that clicking its arrow will also collapse
            // any following paragraph with a lower level, like the paragraphs below
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
            builder.Writeln("Paragraph outline level 2.");

            // Two paragraphs of the same level will not collapse each other
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level3;
            builder.Writeln("Paragraph outline level 3.");
            builder.Writeln("Paragraph outline level 3.");

            // The default "BodyText" value is the lowest
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
            builder.Writeln("Paragraph at main text level.");

            doc.Save(ArtifactsDir + "ParagraphFormat.ParagraphOutlineLevel.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.ParagraphOutlineLevel.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(OutlineLevel.Level1, paragraphs[0].ParagraphFormat.OutlineLevel);
            Assert.AreEqual(OutlineLevel.Level2, paragraphs[1].ParagraphFormat.OutlineLevel);
            Assert.AreEqual(OutlineLevel.Level3, paragraphs[2].ParagraphFormat.OutlineLevel);
            Assert.AreEqual(OutlineLevel.Level3, paragraphs[3].ParagraphFormat.OutlineLevel);
            Assert.AreEqual(OutlineLevel.BodyText, paragraphs[4].ParagraphFormat.OutlineLevel);

        }

        [Test]
        public void PageBreakBefore()
        {
            //ExStart
            //ExFor:ParagraphFormat.PageBreakBefore
            //ExSummary:Shows how to force a page break before each paragraph.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set this to insert a page break before this paragraph
            builder.ParagraphFormat.PageBreakBefore = true;

            // The value we set is propagated to all paragraphs that are created afterwards
            builder.Writeln("Paragraph 1, page 1.");
            builder.Writeln("Paragraph 2, page 2.");

            doc.Save(ArtifactsDir + "ParagraphFormat.PageBreakBefore.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.PageBreakBefore.docx");

            Assert.True(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.PageBreakBefore);
            Assert.True(doc.FirstSection.Body.Paragraphs[1].ParagraphFormat.PageBreakBefore);
        }

        [Test]
        public void WidowControl()
        {
            //ExStart
            //ExFor:ParagraphFormat.WidowControl
            //ExSummary:Shows how to enable widow/orphan control for a paragraph.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text that will not fit on one page, with one line spilling into page 2
            builder.Font.Size = 68;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // This line is referred to as an "Orphan", and a line left behind on the end of the previous page is called a "Widow"
            // They can be fixed by changing size/line spacing/page margins
            // Alternatively, we can use this flag, for which the corresponding Microsoft Word option is 
            // found in Home > Paragraph > Paragraph Settings (button on the bottom right of the tab) 
            // This will add more text to the orphan by putting two lines of text into the second page
            builder.ParagraphFormat.WidowControl = true;

            doc.Save(ArtifactsDir + "ParagraphFormat.WidowControl.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.WidowControl.docx");

            Assert.True(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.WidowControl);
        }

        [Test]
        public void LinesToDrop()
        {
            //ExStart
            //ExFor:ParagraphFormat.LinesToDrop
            //ExSummary:Shows how to set the size of the drop cap text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Setting this attribute will designate the current paragraph as a drop cap,
            // in this case with a height of 4 lines of text
            builder.ParagraphFormat.LinesToDrop = 4;
            builder.Write("H");

            // Any subsequent paragraphs will wrap around the drop cap
            builder.InsertParagraph();
            builder.Write("ello world!");

            doc.Save(ArtifactsDir + "ParagraphFormat.LinesToDrop.odt");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.LinesToDrop.odt");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(4, paragraphs[0].ParagraphFormat.LinesToDrop);
            Assert.AreEqual(0, paragraphs[1].ParagraphFormat.LinesToDrop);
        }
    }
}
