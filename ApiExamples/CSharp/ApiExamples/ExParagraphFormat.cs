// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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

            Assert.AreEqual(OutlineLevel.Level1, paragraphs[0].ParagraphFormat.OutlineLevel);
            Assert.AreEqual(OutlineLevel.Level2, paragraphs[1].ParagraphFormat.OutlineLevel);
            Assert.AreEqual(OutlineLevel.Level3, paragraphs[2].ParagraphFormat.OutlineLevel);
            Assert.AreEqual(OutlineLevel.Level3, paragraphs[3].ParagraphFormat.OutlineLevel);
            Assert.AreEqual(OutlineLevel.BodyText, paragraphs[4].ParagraphFormat.OutlineLevel);

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
                Assert.AreEqual(1, layoutCollector.GetStartPageIndex(paragraphs[0]));
                Assert.AreEqual(2, layoutCollector.GetStartPageIndex(paragraphs[1]));
            }
            else
            {
                Assert.AreEqual(1, layoutCollector.GetStartPageIndex(paragraphs[0]));
                Assert.AreEqual(1, layoutCollector.GetStartPageIndex(paragraphs[1]));
            }

            doc.Save(ArtifactsDir + "ParagraphFormat.PageBreakBefore.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ParagraphFormat.PageBreakBefore.docx");
            paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(pageBreakBefore, paragraphs[0].ParagraphFormat.PageBreakBefore);
            Assert.AreEqual(pageBreakBefore, paragraphs[1].ParagraphFormat.PageBreakBefore);
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

            Assert.AreEqual(widowControl, doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.WidowControl);
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

            Assert.AreEqual(4, paragraphs[0].ParagraphFormat.LinesToDrop);
            Assert.AreEqual(0, paragraphs[1].ParagraphFormat.LinesToDrop);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void SuppressHyphens(bool suppressAutoHyphens)
        {
            //ExStart
            //ExFor:ParagraphFormat.SuppressAutoHyphens
            //ExSummary:Shows how to suppress hyphenation for a paragraph.
            Hyphenation.RegisterDictionary("de-CH", MyDir + "hyph_de_CH.dic");

            Assert.True(Hyphenation.IsDictionaryRegistered("de-CH"));

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

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "ParagraphFormat.SuppressHyphens.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            if (suppressAutoHyphens)
                Assert.True(textAbsorber.Text.Contains("La  ob  storen  an  deinen  am  sachen. \r\n" +
                                                       "Doppelte  um  da  am  spateren  verlogen \r\n" +
                                                       "gekommen  achtzehn  blaulich."));
            else
                Assert.True(textAbsorber.Text.Contains("La ob storen an deinen am sachen. Dop-\r\n" +
                                                       "pelte  um  da  am  spateren  verlogen  ge-\r\n" +
                                                       "kommen  achtzehn  blaulich."));
#endif
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
            Assert.AreEqual(format.LeftIndent, 0.0d);

            format.CharacterUnitLeftIndent = 10.0;

            Assert.AreEqual(format.LeftIndent, 120.0d);

            // 2 -  Right indent:
            Assert.AreEqual(format.RightIndent, 0.0d); 

            format.CharacterUnitRightIndent = -5.5;

            Assert.AreEqual(format.RightIndent, -66.0d);

            // 3 -  Hanging indent:
            Assert.AreEqual(format.FirstLineIndent, 0.0d);

            format.CharacterUnitFirstLineIndent = 20.3;

            Assert.AreEqual(format.FirstLineIndent, 243.59d, 0.1d);

            // 4 -  Line spacing before paragraphs:
            Assert.AreEqual(format.SpaceBefore, 0.0d);

            format.LineUnitBefore = 5.1;

            Assert.AreEqual(format.SpaceBefore, 61.1d, 0.1d);

            // 5 -  Line spacing after paragraphs:
            Assert.AreEqual(format.SpaceAfter, 0.0d);

            format.LineUnitAfter = 10.9;

            Assert.AreEqual(format.SpaceAfter, 130.8d, 0.1d);

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
    }
}
