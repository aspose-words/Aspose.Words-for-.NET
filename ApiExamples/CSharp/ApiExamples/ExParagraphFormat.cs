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

        [TestCase(DropCapPosition.Margin)]
        [TestCase(DropCapPosition.Normal)]
        [TestCase(DropCapPosition.None)]
        public void DropCap(DropCapPosition dropCapPosition)
        {
            //ExStart
            //ExFor:DropCapPosition
            //ExSummary:Shows how to create a drop cap.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert one paragraph with a large letter that the text in the second and third paragraphs begins with.
            builder.Font.Size = 54;
            builder.Writeln("L");

            builder.Font.Size = 18;
            builder.Writeln("orem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
            builder.Writeln("Ut enim ad minim veniam, quis nostrud exercitation " +
                            "ullamco laboris nisi ut aliquip ex ea commodo consequat.");

            // Currently, the second and third paragraphs will appear underneath the first.
            // We can convert the first paragraph as a drop cap for the other paragraphs via its "ParagraphFormat" object.
            // Set the "DropCapPosition" property to "DropCapPosition.Margin" to place the drop cap outside
            // the left hand side page margin, if our text is left-to-right.
            // Set the "DropCapPosition" property to "DropCapPosition.Normal" to place the drop cap within the page margins,
            // and to wrap the rest of the text around it.
            // "DropCapPosition.None" is the default state for all paragraphs.
            ParagraphFormat format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;
            format.DropCapPosition = dropCapPosition;

            doc.Save(ArtifactsDir + "ParagraphFormat.DropCap.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.DropCap.docx");

            Assert.AreEqual(dropCapPosition, doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.DropCapPosition);
            Assert.AreEqual(DropCapPosition.None, doc.FirstSection.Body.Paragraphs[1].ParagraphFormat.DropCapPosition);
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

            // Below are three line spacing rules that we can define using the
            // paragraph's "LineSpacingRule" property to configure spacing between paragraphs.
            // 1 -  Set a minimum amount of spacing.
            // This will give vertical padding to lines of text of any size
            // that's too small to maintain the minimum line height.
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
            builder.ParagraphFormat.LineSpacing = 20;

            builder.Writeln("Minimum line spacing of 20.");
            builder.Writeln("Minimum line spacing of 20.");

            // 2 -  Set exact spacing.
            // Using font sizes that are too large for the spacing will truncate the text.
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            builder.ParagraphFormat.LineSpacing = 5;

            builder.Writeln("Line spacing of exactly 5.");
            builder.Writeln("Line spacing of exactly 5.");

            // 3 -  Set spacing as a multiple of default line spacing, which is 12 points by default.
            // This kind of spacing will scale to different font sizes.
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            builder.ParagraphFormat.LineSpacing = 18;

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

        [TestCase(false)]
        [TestCase(true)]
        public void ParagraphSpacingAuto(bool autoSpacing)
        {
            //ExStart
            //ExFor:ParagraphFormat.SpaceAfter
            //ExFor:ParagraphFormat.SpaceAfterAuto
            //ExFor:ParagraphFormat.SpaceBefore
            //ExFor:ParagraphFormat.SpaceBeforeAuto
            //ExSummary:Shows how to set automatic paragraph spacing.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply a large amount of spacing before and after paragraphs that this builder will create.
            builder.ParagraphFormat.SpaceBefore = 24;
            builder.ParagraphFormat.SpaceAfter = 24;

            // Set these flags to "true" to apply automatic spacing,
            // effectively ignoring the spacing in the attributes we set above.
            // Leave them as "false" will apply our custom paragraph spacing.
            builder.ParagraphFormat.SpaceAfterAuto = autoSpacing;
            builder.ParagraphFormat.SpaceBeforeAuto = autoSpacing;

            // Insert two paragraphs which will have spacing above and below them and save the document.
            builder.Writeln("Paragraph 1.");
            builder.Writeln("Paragraph 2.");

            doc.Save(ArtifactsDir + "ParagraphFormat.ParagraphSpacingAuto.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.ParagraphSpacingAuto.docx");
            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;

            Assert.AreEqual(24.0d, format.SpaceBefore);
            Assert.AreEqual(24.0d, format.SpaceAfter);
            Assert.AreEqual(autoSpacing, format.SpaceAfterAuto);
            Assert.AreEqual(autoSpacing, format.SpaceBeforeAuto);

            format = doc.FirstSection.Body.Paragraphs[1].ParagraphFormat;

            Assert.AreEqual(24.0d, format.SpaceBefore);
            Assert.AreEqual(24.0d, format.SpaceAfter);
            Assert.AreEqual(autoSpacing, format.SpaceAfterAuto);
            Assert.AreEqual(autoSpacing, format.SpaceBeforeAuto);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ParagraphSpacingSameStyle(bool noSpaceBetweenParagraphsOfSameStyle)
        {
            //ExStart
            //ExFor:ParagraphFormat.SpaceAfter
            //ExFor:ParagraphFormat.SpaceBefore
            //ExFor:ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle
            //ExSummary:Shows how to set automatic paragraph spacing.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply a large amount of spacing before and after paragraphs that this builder will create.
            builder.ParagraphFormat.SpaceBefore = 24;
            builder.ParagraphFormat.SpaceAfter = 24;

            // Set this flag to "true" to apply no spacing between paragraphs
            // with the same style, which will group similar paragraphs together.
            // Leave ths flag as "false" to evenly apply spacing to every paragraph.
            builder.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle = noSpaceBetweenParagraphsOfSameStyle;

            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln($"Paragraph in the \"{builder.ParagraphFormat.Style.Name}\" style.");
            builder.Writeln($"Paragraph in the \"{builder.ParagraphFormat.Style.Name}\" style.");
            builder.Writeln($"Paragraph in the \"{builder.ParagraphFormat.Style.Name}\" style.");
            builder.ParagraphFormat.Style = doc.Styles["Quote"];
            builder.Writeln($"Paragraph in the \"{builder.ParagraphFormat.Style.Name}\" style.");
            builder.Writeln($"Paragraph in the \"{builder.ParagraphFormat.Style.Name}\" style.");
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln($"Paragraph in the \"{builder.ParagraphFormat.Style.Name}\" style.");
            builder.Writeln($"Paragraph in the \"{builder.ParagraphFormat.Style.Name}\" style.");

            doc.Save(ArtifactsDir + "ParagraphFormat.ParagraphSpacingSameStyle.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.ParagraphSpacingSameStyle.docx");

            foreach (Paragraph paragraph in doc.FirstSection.Body.Paragraphs)
            {
                ParagraphFormat format = paragraph.ParagraphFormat;

                Assert.AreEqual(24.0d, format.SpaceBefore);
                Assert.AreEqual(24.0d, format.SpaceAfter);
                Assert.AreEqual(noSpaceBetweenParagraphsOfSameStyle, format.NoSpaceBetweenParagraphsOfSameStyle);
            }
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

        [Test]
        public void SuppressHyphens()
        {
            //ExStart
            //ExFor:ParagraphFormat.SuppressAutoHyphens
            //ExSummary:Shows how to suppress document hyphenation.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Size = 24;
            builder.ParagraphFormat.SuppressAutoHyphens = false;

            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            doc.Save(ArtifactsDir + "ParagraphFormat.SuppressHyphens.docx");
            //ExEnd
        }
        
        public void ParagraphSpacingAndIndents()
        {
            //ExStart
            //ExFor:ParagraphFormat.CharacterUnitLeftIndent
            //ExFor:ParagraphFormat.CharacterUnitRightIndent
            //ExFor:ParagraphFormat.CharacterUnitFirstLineIndent
            //ExFor:ParagraphFormat.LineUnitBefore
            //ExFor:ParagraphFormat.LineUnitAfter
            //ExSummary:Shows how to change paragraph spacing and indents.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            ParagraphFormat format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;
            
            Assert.AreEqual(format.LeftIndent, 0.0d); //ExSkip
            Assert.AreEqual(format.RightIndent, 0.0d); //ExSkip
            Assert.AreEqual(format.FirstLineIndent, 0.0d); //ExSkip
            Assert.AreEqual(format.SpaceBefore, 0.0d); //ExSkip
            Assert.AreEqual(format.SpaceAfter, 0.0d); //ExSkip

            // Also ParagraphFormat.LeftIndent will be updated
            format.CharacterUnitLeftIndent = 10.0;
            // Also ParagraphFormat.RightIndent will be updated
            format.CharacterUnitRightIndent = -5.5;
            // Also ParagraphFormat.FirstLineIndent will be updated
            format.CharacterUnitFirstLineIndent = 20.3;
            // Also ParagraphFormat.SpaceBefore will be updated
            format.LineUnitBefore = 5.1;
            // Also ParagraphFormat.SpaceAfter will be updated
            format.LineUnitAfter= 10.9;

            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.Write("测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试" +
                          "文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档");
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;
            
            Assert.AreEqual(format.CharacterUnitLeftIndent, 10.0d);
            Assert.AreEqual(format.LeftIndent, 120.0d);
            
            Assert.AreEqual(format.CharacterUnitRightIndent, -5.5d);
            Assert.AreEqual(format.RightIndent, -66.0d);
            
            Assert.AreEqual(format.CharacterUnitFirstLineIndent, 20.3d);
            Assert.AreEqual(format.FirstLineIndent, 243.59d, 0.1d);
            
            Assert.AreEqual(format.LineUnitBefore, 5.1d, 0.1d);
            Assert.AreEqual(format.SpaceBefore, 61.1d, 0.1d);
            
            Assert.AreEqual(format.LineUnitAfter, 10.9d);
            Assert.AreEqual(format.SpaceAfter, 130.8d, 0.1d);
        }

        [Test]
        public void SnapToGrid()
        {
            //ExStart
            //ExFor:ParagraphFormat.SnapToGrid
            //ExSummary:Shows how to work with extremely wide spacing in the document.
            Document doc = new Document();
            Paragraph par = doc.FirstSection.Body.FirstParagraph;
            // Set 'SnapToGrid' to true if need optimize the layout when typing in Asian characters
            // Use 'SnapToGrid' for the whole paragraph
            par.ParagraphFormat.SnapToGrid = true;
            
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod " +
                            "tempor incididunt ut labore et dolore magna aliqua.");
            // Use 'SnapToGrid' for the specific run
            par.Runs[0].Font.SnapToGrid = true;

            doc.Save(ArtifactsDir + "Paragraph.SnapToGrid.docx");
        }
    }
}
