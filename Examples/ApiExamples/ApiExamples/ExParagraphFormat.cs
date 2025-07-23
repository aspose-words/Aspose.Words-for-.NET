// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Pdf.Text;
using Aspose.Words;
using Aspose.Words.Layout;
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

            Assert.That(format.FarEastLineBreakControl, Is.True);
            Assert.That(format.WordWrap, Is.False);
            Assert.That(format.HangingPunctuation, Is.True);
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
            // Set the "DropCapPosition" property to "DropCapPosition.Margin" to place the drop cap
            // outside the left-hand side page margin if our text is left-to-right.
            // Set the "DropCapPosition" property to "DropCapPosition.Normal" to place the drop cap within the page margins
            // and to wrap the rest of the text around it.
            // "DropCapPosition.None" is the default state for all paragraphs.
            ParagraphFormat format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;
            format.DropCapPosition = dropCapPosition;

            doc.Save(ArtifactsDir + "ParagraphFormat.DropCap.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.DropCap.docx");

            Assert.That(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.DropCapPosition, Is.EqualTo(dropCapPosition));
            Assert.That(doc.FirstSection.Body.Paragraphs[1].ParagraphFormat.DropCapPosition, Is.EqualTo(DropCapPosition.None));
        }

        [Test]
        public void LineSpacing()
        {
            //ExStart
            //ExFor:ParagraphFormat.LineSpacing
            //ExFor:ParagraphFormat.LineSpacingRule
            //ExFor:LineSpacingRule
            //ExSummary:Shows how to work with line spacing.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are three line spacing rules that we can define using the
            // paragraph's "LineSpacingRule" property to configure spacing between paragraphs.
            // 1 -  Set a minimum amount of spacing.
            // This will give vertical padding to lines of text of any size
            // that is too small to maintain the minimum line-height.
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

            Assert.That(paragraphs[0].ParagraphFormat.LineSpacingRule, Is.EqualTo(LineSpacingRule.AtLeast));
            Assert.That(paragraphs[0].ParagraphFormat.LineSpacing, Is.EqualTo(20.0d));
            Assert.That(paragraphs[1].ParagraphFormat.LineSpacingRule, Is.EqualTo(LineSpacingRule.AtLeast));
            Assert.That(paragraphs[1].ParagraphFormat.LineSpacing, Is.EqualTo(20.0d));

            Assert.That(paragraphs[2].ParagraphFormat.LineSpacingRule, Is.EqualTo(LineSpacingRule.Exactly));
            Assert.That(paragraphs[2].ParagraphFormat.LineSpacing, Is.EqualTo(5.0d));
            Assert.That(paragraphs[3].ParagraphFormat.LineSpacingRule, Is.EqualTo(LineSpacingRule.Exactly));
            Assert.That(paragraphs[3].ParagraphFormat.LineSpacing, Is.EqualTo(5.0d));

            Assert.That(paragraphs[4].ParagraphFormat.LineSpacingRule, Is.EqualTo(LineSpacingRule.Multiple));
            Assert.That(paragraphs[4].ParagraphFormat.LineSpacing, Is.EqualTo(18.0d));
            Assert.That(paragraphs[5].ParagraphFormat.LineSpacingRule, Is.EqualTo(LineSpacingRule.Multiple));
            Assert.That(paragraphs[5].ParagraphFormat.LineSpacing, Is.EqualTo(18.0d));
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
            // effectively ignoring the spacing in the properties we set above.
            // Leave them as "false" will apply our custom paragraph spacing.
            builder.ParagraphFormat.SpaceAfterAuto = autoSpacing;
            builder.ParagraphFormat.SpaceBeforeAuto = autoSpacing;

            // Insert two paragraphs that will have spacing above and below them and save the document.
            builder.Writeln("Paragraph 1.");
            builder.Writeln("Paragraph 2.");

            doc.Save(ArtifactsDir + "ParagraphFormat.ParagraphSpacingAuto.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.ParagraphSpacingAuto.docx");
            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;

            Assert.That(format.SpaceBefore, Is.EqualTo(24.0d));
            Assert.That(format.SpaceAfter, Is.EqualTo(24.0d));
            Assert.That(format.SpaceAfterAuto, Is.EqualTo(autoSpacing));
            Assert.That(format.SpaceBeforeAuto, Is.EqualTo(autoSpacing));

            format = doc.FirstSection.Body.Paragraphs[1].ParagraphFormat;

            Assert.That(format.SpaceBefore, Is.EqualTo(24.0d));
            Assert.That(format.SpaceAfter, Is.EqualTo(24.0d));
            Assert.That(format.SpaceAfterAuto, Is.EqualTo(autoSpacing));
            Assert.That(format.SpaceBeforeAuto, Is.EqualTo(autoSpacing));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ParagraphSpacingSameStyle(bool noSpaceBetweenParagraphsOfSameStyle)
        {
            //ExStart
            //ExFor:ParagraphFormat.SpaceAfter
            //ExFor:ParagraphFormat.SpaceBefore
            //ExFor:ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle
            //ExSummary:Shows how to apply no spacing between paragraphs with the same style.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply a large amount of spacing before and after paragraphs that this builder will create.
            builder.ParagraphFormat.SpaceBefore = 24;
            builder.ParagraphFormat.SpaceAfter = 24;

            // Set the "NoSpaceBetweenParagraphsOfSameStyle" flag to "true" to apply
            // no spacing between paragraphs with the same style, which will group similar paragraphs.
            // Leave the "NoSpaceBetweenParagraphsOfSameStyle" flag as "false"
            // to evenly apply spacing to every paragraph.
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

                Assert.That(format.SpaceBefore, Is.EqualTo(24.0d));
                Assert.That(format.SpaceAfter, Is.EqualTo(24.0d));
                Assert.That(format.NoSpaceBetweenParagraphsOfSameStyle, Is.EqualTo(noSpaceBetweenParagraphsOfSameStyle));
            }
        }

        [Test]
        public void ParagraphOutlineLevel()
        {
            //ExStart
            //ExFor:ParagraphFormat.OutlineLevel
            //ExFor:OutlineLevel
            //ExSummary:Shows how to configure paragraph outline levels to create collapsible text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Each paragraph has an OutlineLevel, which could be any number from 1 to 9, or at the default "BodyText" value.
            // Setting the property to one of the numbered values will show an arrow to the left
            // of the beginning of the paragraph.
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
            builder.Writeln("Paragraph outline level 1.");

            // Level 1 is the topmost level. If there is a paragraph with a lower level below a paragraph with a higher level,
            // collapsing the higher-level paragraph will collapse the lower level paragraph.
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
            builder.Writeln("Paragraph outline level 2.");

            // Two paragraphs of the same level will not collapse each other,
            // and the arrows do not collapse the paragraphs they point to.
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.Level3;
            builder.Writeln("Paragraph outline level 3.");
            builder.Writeln("Paragraph outline level 3.");

            // The default "BodyText" value is the lowest, which a paragraph of any level can collapse.
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
            builder.Writeln("Paragraph at main text level.");

            doc.Save(ArtifactsDir + "ParagraphFormat.ParagraphOutlineLevel.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.ParagraphOutlineLevel.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.That(paragraphs[0].ParagraphFormat.OutlineLevel, Is.EqualTo(OutlineLevel.Level1));
            Assert.That(paragraphs[1].ParagraphFormat.OutlineLevel, Is.EqualTo(OutlineLevel.Level2));
            Assert.That(paragraphs[2].ParagraphFormat.OutlineLevel, Is.EqualTo(OutlineLevel.Level3));
            Assert.That(paragraphs[3].ParagraphFormat.OutlineLevel, Is.EqualTo(OutlineLevel.Level3));
            Assert.That(paragraphs[4].ParagraphFormat.OutlineLevel, Is.EqualTo(OutlineLevel.BodyText));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void PageBreakBefore(bool pageBreakBefore)
        {
            //ExStart
            //ExFor:ParagraphFormat.PageBreakBefore
            //ExSummary:Shows how to create paragraphs with page breaks at the beginning.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set this flag to "true" to apply a page break to each paragraph's beginning
            // that the document builder will create under this ParagraphFormat configuration.
            // The first paragraph will not receive a page break.
            // Leave this flag as "false" to start each new paragraph on the same page
            // as the previous, provided there is sufficient space.
            builder.ParagraphFormat.PageBreakBefore = pageBreakBefore;

            builder.Writeln("Paragraph 1.");
            builder.Writeln("Paragraph 2.");

            LayoutCollector layoutCollector = new LayoutCollector(doc);
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            
            if (pageBreakBefore)
            {
                Assert.That(layoutCollector.GetStartPageIndex(paragraphs[0]), Is.EqualTo(1));
                Assert.That(layoutCollector.GetStartPageIndex(paragraphs[1]), Is.EqualTo(2));
            }
            else
            {
                Assert.That(layoutCollector.GetStartPageIndex(paragraphs[0]), Is.EqualTo(1));
                Assert.That(layoutCollector.GetStartPageIndex(paragraphs[1]), Is.EqualTo(1));
            }

            doc.Save(ArtifactsDir + "ParagraphFormat.PageBreakBefore.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.PageBreakBefore.docx");
            paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.That(paragraphs[0].ParagraphFormat.PageBreakBefore, Is.EqualTo(pageBreakBefore));
            Assert.That(paragraphs[1].ParagraphFormat.PageBreakBefore, Is.EqualTo(pageBreakBefore));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void WidowControl(bool widowControl)
        {
            //ExStart
            //ExFor:ParagraphFormat.WidowControl
            //ExSummary:Shows how to enable widow/orphan control for a paragraph.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // When we write the text that does not fit onto one page, one line may spill over onto the next page.
            // The single line that ends up on the next page is called an "Orphan",
            // and the previous line where the orphan broke off is called a "Widow".
            // We can fix orphans and widows by rearranging text via font size, spacing, or page margins.
            // If we wish to preserve our document's dimensions, we can set this flag to "true"
            // to push widows onto the same page as their respective orphans. 
            // Leave this flag as "false" will leave widow/orphan pairs in text.
            // Every paragraph has this setting accessible in Microsoft Word via Home -> Paragraph -> Paragraph Settings
            // (button on bottom right hand corner of "Paragraph" tab) -> "Widow/Orphan control".
            builder.ParagraphFormat.WidowControl = widowControl; 

            // Insert text that produces an orphan and a widow.
            builder.Font.Size = 68;
            builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            doc.Save(ArtifactsDir + "ParagraphFormat.WidowControl.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.WidowControl.docx");

            Assert.That(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.WidowControl, Is.EqualTo(widowControl));
        }

        [Test]
        public void LinesToDrop()
        {
            //ExStart
            //ExFor:ParagraphFormat.LinesToDrop
            //ExSummary:Shows how to set the size of a drop cap.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Modify the "LinesToDrop" property to designate a paragraph as a drop cap,
            // which will turn it into a large capital letter that will decorate the next paragraph.
            // Give this property a value of 4 to give the drop cap the height of four text lines.
            builder.ParagraphFormat.LinesToDrop = 4;
            builder.Writeln("H");

            // Reset the "LinesToDrop" property to 0 to turn the next paragraph into an ordinary paragraph.
            // The text in this paragraph will wrap around the drop cap.
            builder.ParagraphFormat.LinesToDrop = 0;
            builder.Writeln("ello world!");

            doc.Save(ArtifactsDir + "ParagraphFormat.LinesToDrop.odt");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.LinesToDrop.odt");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.That(paragraphs[0].ParagraphFormat.LinesToDrop, Is.EqualTo(4));
            Assert.That(paragraphs[1].ParagraphFormat.LinesToDrop, Is.EqualTo(0));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void SuppressHyphens(bool suppressAutoHyphens)
        {
            //ExStart
            //ExFor:ParagraphFormat.SuppressAutoHyphens
            //ExSummary:Shows how to suppress hyphenation for a paragraph.
            Hyphenation.RegisterDictionary("de-CH", MyDir + "hyph_de_CH.dic");

            Assert.That(Hyphenation.IsDictionaryRegistered("de-CH"), Is.True);

            // Open a document containing text with a locale matching that of our dictionary.
            // When we save this document to a fixed page save format, its text will have hyphenation.
            Document doc = new Document(MyDir + "German text.docx");

            // We can set the "SuppressAutoHyphens" property to "true" to disable hyphenation
            // for a specific paragraph while keeping it enabled for the rest of the document.
            // The default value for this property is "false",
            // which means every paragraph by default uses hyphenation if any is available.
            doc.FirstSection.Body.FirstParagraph.ParagraphFormat.SuppressAutoHyphens = suppressAutoHyphens;

            doc.Save(ArtifactsDir + "ParagraphFormat.SuppressHyphens.pdf");
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void UsePdfDocumentForSuppressHyphens(bool suppressAutoHyphens)
        {
            const string unicodeOptionalHyphen = "\xad";

            SuppressHyphens(suppressAutoHyphens);

            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "ParagraphFormat.SuppressHyphens.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            if (suppressAutoHyphens)
                Assert.That(textAbsorber.Text.Replace("  ", " ").Contains($"La ob storen an deinen am sachen. {Environment.NewLine}" +
                                                       $"Doppelte um da am spateren verlogen {Environment.NewLine}" +
                                                       $"gekommen achtzehn blaulich."), Is.True);
            else
                Assert.That(textAbsorber.Text.Replace("  ", " ").Contains($"La ob storen an deinen am sachen. Dop{unicodeOptionalHyphen}{Environment.NewLine}" +
                                                       $"pelte um da am spateren verlogen ge{unicodeOptionalHyphen}{Environment.NewLine}" +
                                                       $"kommen achtzehn blaulich."), Is.True);
        }

        [Test]
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

            // Below are five different spacing options, along with the properties that their configuration indirectly affects.
            // 1 -  Left indent:
            Assert.That(0.0d, Is.EqualTo(format.LeftIndent));

            format.CharacterUnitLeftIndent = 10.0;

            Assert.That(120.0d, Is.EqualTo(format.LeftIndent));

            // 2 -  Right indent:
            Assert.That(0.0d, Is.EqualTo(format.RightIndent)); 

            format.CharacterUnitRightIndent = -5.5;

            Assert.That(-66.0d, Is.EqualTo(format.RightIndent));

            // 3 -  Hanging indent:
            Assert.That(0.0d, Is.EqualTo(format.FirstLineIndent));

            format.CharacterUnitFirstLineIndent = 20.3;

            Assert.That(243.59d, Is.EqualTo(format.FirstLineIndent).Within(0.1d));

            // 4 -  Line spacing before paragraphs:
            Assert.That(0.0d, Is.EqualTo(format.SpaceBefore));

            format.LineUnitBefore = 5.1;

            Assert.That(61.1d, Is.EqualTo(format.SpaceBefore).Within(0.1d));

            // 5 -  Line spacing after paragraphs:
            Assert.That(0.0d, Is.EqualTo(format.SpaceAfter));

            format.LineUnitAfter = 10.9;

            Assert.That(130.8d, Is.EqualTo(format.SpaceAfter).Within(0.1d));

            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.Write("测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试" +
                          "文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档");
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;

            Assert.That(10.0d, Is.EqualTo(format.CharacterUnitLeftIndent));
            Assert.That(120.0d, Is.EqualTo(format.LeftIndent));
            
            Assert.That(-5.5d, Is.EqualTo(format.CharacterUnitRightIndent));
            Assert.That(-66.0d, Is.EqualTo(format.RightIndent));

            Assert.That(20.3d, Is.EqualTo(format.CharacterUnitFirstLineIndent));
            Assert.That(243.59d, Is.EqualTo(format.FirstLineIndent).Within(0.1d));
            
            Assert.That(5.1d, Is.EqualTo(format.LineUnitBefore).Within(0.1d));
            Assert.That(61.1d, Is.EqualTo(format.SpaceBefore).Within(0.1d));

            Assert.That(10.9d, Is.EqualTo(format.LineUnitAfter));
            Assert.That(130.8d, Is.EqualTo(format.SpaceAfter).Within(0.1d));
        }

        [Test]
        public void ParagraphBaselineAlignment()
        {
            //ExStart
            //ExFor:BaselineAlignment
            //ExFor:ParagraphFormat.BaselineAlignment
            //ExSummary:Shows how to set fonts vertical position on a line.
            Document doc = new Document(MyDir + "Office math.docx");

            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;
            if (format.BaselineAlignment == BaselineAlignment.Auto)
            {                
                format.BaselineAlignment = BaselineAlignment.Top;
            }

            doc.Save(ArtifactsDir + "ParagraphFormat.ParagraphBaselineAlignment.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.ParagraphBaselineAlignment.docx");
            format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;
            Assert.That(format.BaselineAlignment, Is.EqualTo(BaselineAlignment.Top));
        }

        [Test]
        public void MirrorIndents()
        {
            //ExStart:MirrorIndents
            //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
            //ExFor:ParagraphFormat.MirrorIndents
            //ExSummary:Show how to make left and right indents the same.
            Document doc = new Document(MyDir + "Document.docx");
            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;

            format.MirrorIndents = true;

            doc.Save(ArtifactsDir + "ParagraphFormat.MirrorIndents.docx");
            //ExEnd:MirrorIndents

            doc = new Document(ArtifactsDir + "ParagraphFormat.MirrorIndents.docx");
            format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;

            Assert.That(format.MirrorIndents, Is.EqualTo(true));
        }
    }
}
