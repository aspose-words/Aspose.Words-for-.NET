// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using Aspose.Words;
using Aspose.Words.Settings;
using NUnit.Framework;
using PaperSize = Aspose.Words.PaperSize;
#if NET462 || NETCOREAPP2_1 || JAVA
using System.Drawing.Printing;
using System.Linq;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExPageSetup : ApiExampleBase
    {
        [Test]
        public void ClearFormatting()
        {
            //ExStart
            //ExFor:DocumentBuilder.PageSetup
            //ExFor:DocumentBuilder.InsertBreak
            //ExFor:DocumentBuilder.Document
            //ExFor:PageSetup
            //ExFor:PageSetup.Orientation
            //ExFor:PageSetup.VerticalAlignment
            //ExFor:PageSetup.ClearFormatting
            //ExFor:Orientation
            //ExFor:PageVerticalAlignment
            //ExFor:BreakType
            //ExSummary:Shows how to insert sections using DocumentBuilder, specify page setup for a section and reset page setup to defaults.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Modify the first section in the document
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.VerticalAlignment = PageVerticalAlignment.Center;
            builder.Writeln("Section 1, landscape oriented and text vertically centered.");

            // Start a new section and reset its formatting to defaults
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.PageSetup.ClearFormatting();
            builder.Writeln("Section 2, back to default Letter paper size, portrait orientation and top alignment.");

            doc.Save(ArtifactsDir + "PageSetup.ClearFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.ClearFormatting.docx");

            Assert.AreEqual(Orientation.Landscape, doc.Sections[0].PageSetup.Orientation);
            Assert.AreEqual(PageVerticalAlignment.Center, doc.Sections[0].PageSetup.VerticalAlignment);

            Assert.AreEqual(Orientation.Portrait, doc.Sections[1].PageSetup.Orientation);
            Assert.AreEqual(PageVerticalAlignment.Top, doc.Sections[1].PageSetup.VerticalAlignment);
        }

        [Test]
        public void DifferentHeaders()
        {
            //ExStart
            //ExFor:PageSetup.DifferentFirstPageHeaderFooter
            //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
            //ExFor:PageSetup.LayoutMode
            //ExFor:PageSetup.CharactersPerLine
            //ExFor:PageSetup.LinesPerPage
            //ExFor:SectionLayoutMode
            //ExSummary:Shows how to create headers and footers different for first, even and odd pages using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetup = builder.PageSetup;
            pageSetup.DifferentFirstPageHeaderFooter = true;
            pageSetup.OddAndEvenPagesHeaderFooter = true;
            pageSetup.LayoutMode = SectionLayoutMode.LineGrid;
            pageSetup.CharactersPerLine = 1;
            pageSetup.LinesPerPage = 1;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Writeln("First page header.");

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Writeln("Even pages header.");

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Odd pages header.");

            // Move back to the main story of the first section
            builder.MoveToSection(0);
            builder.Writeln("Text page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Text page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Text page 3.");

            doc.Save(ArtifactsDir + "PageSetup.DifferentHeaders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.DifferentHeaders.docx");

            Assert.True(pageSetup.DifferentFirstPageHeaderFooter);
            Assert.True(pageSetup.OddAndEvenPagesHeaderFooter);
            Assert.AreEqual(SectionLayoutMode.LineGrid, doc.FirstSection.PageSetup.LayoutMode);
            Assert.AreEqual(1, doc.FirstSection.PageSetup.CharactersPerLine);
            Assert.AreEqual(1, doc.FirstSection.PageSetup.LinesPerPage);
        }

        [Test]
        public void SetSectionStart()
        {
            //ExStart
            //ExFor:SectionStart
            //ExFor:PageSetup.SectionStart
            //ExFor:Document.Sections
            //ExSummary:Shows how to specify how the section starts, from a new page, on the same page or other.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add text to the first section and that comes with a blank document,
            // then add a new section that starts a new page and give it text as well
            builder.Writeln("This text is in section 1.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("This text is in section 2.");

            // Section break types determine how a new section gets split from the previous section
            // By inserting a "SectionBreakNewPage" type section break, we've set this section's SectionStart value to "NewPage" 
            Assert.AreEqual(SectionStart.NewPage, doc.Sections[1].PageSetup.SectionStart);

            // Insert a new column section the same way
            builder.InsertBreak(BreakType.SectionBreakNewColumn);
            builder.Writeln("This text is in section 3.");

            Assert.AreEqual(SectionStart.NewColumn, doc.Sections[2].PageSetup.SectionStart);

            // We can change the types of section breaks by assigning different values to each section's SectionStart
            // Setting their values to "Continuous" will put no visible breaks between sections
            // and will leave all the content of this document on one page
            doc.Sections[1].PageSetup.SectionStart = SectionStart.Continuous;
            doc.Sections[2].PageSetup.SectionStart = SectionStart.Continuous;

            doc.Save(ArtifactsDir + "PageSetup.SetSectionStart.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.SetSectionStart.docx");

            Assert.AreEqual(SectionStart.NewPage, doc.Sections[0].PageSetup.SectionStart);
            Assert.AreEqual(SectionStart.Continuous, doc.Sections[1].PageSetup.SectionStart);
            Assert.AreEqual(SectionStart.Continuous, doc.Sections[2].PageSetup.SectionStart);
        }

#if NET462 || NETCOREAPP2_1 || JAVA
        [Test]
        [Ignore("Run only when the printer driver is installed")]
        public void DefaultPaperTray()
        {
            //ExStart
            //ExFor:PageSetup.FirstPageTray
            //ExFor:PageSetup.OtherPagesTray
            //ExSummary:Shows how to change all sections in a document to use the default paper tray of the selected printer.
            Document doc = new Document();

            // Find the printer that will be used for printing this document
            // In this case it is the default printer
            // You can define a specific printer using PrinterName
            PrinterSettings settings = new PrinterSettings();

            // The paper tray value stored in documents is completely printer specific
            // This means the code below resets all page tray values to use the current printers default tray
            // You can enumerate PrinterSettings.PaperSources to find the other valid paper tray values of the selected printer
            foreach (Section section in doc.Sections.OfType<Section>())
            {
                section.PageSetup.FirstPageTray = settings.DefaultPageSettings.PaperSource.RawKind;
                section.PageSetup.OtherPagesTray = settings.DefaultPageSettings.PaperSource.RawKind;
            }
            //ExEnd
            
            foreach (Section section in DocumentHelper.SaveOpen(doc).Sections.OfType<Section>())
            {
                Assert.AreEqual(settings.DefaultPageSettings.PaperSource.RawKind, section.PageSetup.FirstPageTray);
                Assert.AreEqual(settings.DefaultPageSettings.PaperSource.RawKind, section.PageSetup.OtherPagesTray);
            }
        }

        [Test]
        [Ignore("Run only when the printer driver is installed")]
        public void PaperTrayForDifferentPaperType()
        {
            //ExStart
            //ExFor:PageSetup.FirstPageTray
            //ExFor:PageSetup.OtherPagesTray
            //ExSummary:Shows how to set up printing using different printer trays for different paper sizes.
            Document doc = new Document();

            // Choose the default printer to be used for printing this document
            PrinterSettings settings = new PrinterSettings();

            // This is the tray we will use for A4 paper size
            // This is the first tray in the paper sources collection
            int printerTrayForA4 = settings.PaperSources[0].RawKind;
            // The is the tray we will use for Letter paper size
            // This is the second tray in the paper sources collection
            int printerTrayForLetter = settings.PaperSources[1].RawKind;

            // Set the page tray used for each section based off the paper size used in the section
            foreach (Section section in doc.Sections.OfType<Section>())
            {
                if (section.PageSetup.PaperSize == Aspose.Words.PaperSize.Letter)
                {
                    section.PageSetup.FirstPageTray = printerTrayForLetter;
                    section.PageSetup.OtherPagesTray = printerTrayForLetter;
                }
                else if (section.PageSetup.PaperSize == Aspose.Words.PaperSize.A4)
                {
                    section.PageSetup.FirstPageTray = printerTrayForA4;
                    section.PageSetup.OtherPagesTray = printerTrayForA4;
                }
            }
            //ExEnd

            foreach (Section section in DocumentHelper.SaveOpen(doc).Sections.OfType<Section>())
            {
                if (section.PageSetup.PaperSize == Aspose.Words.PaperSize.Letter)
                {
                    Assert.AreEqual(printerTrayForLetter, section.PageSetup.FirstPageTray);
                    Assert.AreEqual(printerTrayForLetter, section.PageSetup.OtherPagesTray);
                }
                else if (section.PageSetup.PaperSize == Aspose.Words.PaperSize.A4)
                {
                    Assert.AreEqual(printerTrayForA4, section.PageSetup.FirstPageTray);
                    Assert.AreEqual(printerTrayForA4, section.PageSetup.OtherPagesTray);
                }
            }
        }
#endif

        [Test]
        public void PageMargins()
        {
            //ExStart
            //ExFor:ConvertUtil
            //ExFor:ConvertUtil.InchToPoint
            //ExFor:PaperSize
            //ExFor:PageSetup.PaperSize
            //ExFor:PageSetup.Orientation
            //ExFor:PageSetup.TopMargin
            //ExFor:PageSetup.BottomMargin
            //ExFor:PageSetup.LeftMargin
            //ExFor:PageSetup.RightMargin
            //ExFor:PageSetup.HeaderDistance
            //ExFor:PageSetup.FooterDistance
            //ExSummary:Shows how to adjust paper size, orientation, margins and other settings for a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.PageSetup.PaperSize = PaperSize.Legal;
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.TopMargin = ConvertUtil.InchToPoint(1.0);
            builder.PageSetup.BottomMargin = ConvertUtil.InchToPoint(1.0);
            builder.PageSetup.LeftMargin = ConvertUtil.InchToPoint(1.5);
            builder.PageSetup.RightMargin = ConvertUtil.InchToPoint(1.5);
            builder.PageSetup.HeaderDistance = ConvertUtil.InchToPoint(0.2);
            builder.PageSetup.FooterDistance = ConvertUtil.InchToPoint(0.2);

            builder.Writeln("Hello world.");

            doc.Save(ArtifactsDir + "PageSetup.PageMargins.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PageMargins.docx");

            Assert.AreEqual(PaperSize.Legal, doc.FirstSection.PageSetup.PaperSize);
            Assert.AreEqual(Orientation.Landscape, doc.FirstSection.PageSetup.Orientation);
            Assert.AreEqual(72.0d, doc.FirstSection.PageSetup.TopMargin);
            Assert.AreEqual(72.0d, doc.FirstSection.PageSetup.BottomMargin);
            Assert.AreEqual(108.0d, doc.FirstSection.PageSetup.LeftMargin);
            Assert.AreEqual(108.0d, doc.FirstSection.PageSetup.RightMargin);
            Assert.AreEqual(14.4d, doc.FirstSection.PageSetup.HeaderDistance);
            Assert.AreEqual(14.4d, doc.FirstSection.PageSetup.FooterDistance);
        }

        [Test]
        public void ColumnsSameWidth()
        {
            //ExStart
            //ExFor:PageSetup.TextColumns
            //ExFor:TextColumnCollection
            //ExFor:TextColumnCollection.Spacing
            //ExFor:TextColumnCollection.SetCount
            //ExSummary:Shows how to create multiple evenly spaced columns in a section using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            TextColumnCollection columns = builder.PageSetup.TextColumns;
            // Make spacing between columns wider
            columns.Spacing = 100;
            // This creates two columns of equal width
            columns.SetCount(2);

            builder.Writeln("Text in column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Text in column 2.");

            doc.Save(ArtifactsDir + "PageSetup.ColumnsSameWidth.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.ColumnsSameWidth.docx");

            Assert.AreEqual(100.0d, doc.FirstSection.PageSetup.TextColumns.Spacing);
            Assert.AreEqual(2, doc.FirstSection.PageSetup.TextColumns.Count);
        }

        [Test]
        public void CustomColumnWidth()
        {
            //ExStart
            //ExFor:TextColumnCollection.LineBetween
            //ExFor:TextColumnCollection.EvenlySpaced
            //ExFor:TextColumnCollection.Item
            //ExFor:TextColumn
            //ExFor:TextColumn.Width
            //ExFor:TextColumn.SpaceAfter
            //ExSummary:Shows how to set widths of columns.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            TextColumnCollection columns = builder.PageSetup.TextColumns;
            // Show vertical line between columns
            columns.LineBetween = true;
            // Indicate we want to create column with different widths
            columns.EvenlySpaced = false;
            // Create two columns, note they will be created with zero widths, need to set them
            columns.SetCount(2);

            // Set the first column to be narrow
            TextColumn column = columns[0];
            column.Width = 100;
            column.SpaceAfter = 20;

            // Set the second column to take the rest of the space available on the page
            column = columns[1];
            PageSetup pageSetup = builder.PageSetup;
            double contentWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
            column.Width = contentWidth - column.Width - column.SpaceAfter;

            builder.Writeln("Narrow column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Wide column 2.");

            doc.Save(ArtifactsDir + "PageSetup.CustomColumnWidth.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.CustomColumnWidth.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.True(pageSetup.TextColumns.LineBetween);
            Assert.False(pageSetup.TextColumns.EvenlySpaced);
            Assert.AreEqual(2, pageSetup.TextColumns.Count);
            Assert.AreEqual(100.0d, pageSetup.TextColumns[0].Width);
            Assert.AreEqual(20.0d, pageSetup.TextColumns[0].SpaceAfter);
            Assert.AreEqual(468.0d, pageSetup.TextColumns[1].Width);
            Assert.AreEqual(0.0d, pageSetup.TextColumns[1].SpaceAfter);
        }

        [Test]
        public void LineNumbers()
        {
            //ExStart
            //ExFor:PageSetup.LineStartingNumber
            //ExFor:PageSetup.LineNumberDistanceFromText
            //ExFor:PageSetup.LineNumberCountBy
            //ExFor:PageSetup.LineNumberRestartMode
            //ExFor:ParagraphFormat.SuppressLineNumbers
            //ExFor:LineNumberRestartMode
            //ExSummary:Shows how to enable Microsoft Word line numbering for a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Line numbering for each section can be configured via PageSetup
            PageSetup pageSetup = builder.PageSetup;
            pageSetup.LineStartingNumber = 1;
            pageSetup.LineNumberCountBy = 3;
            pageSetup.LineNumberRestartMode = LineNumberRestartMode.RestartPage;
            pageSetup.LineNumberDistanceFromText = 50.0d;

            // LineNumberCountBy is set to 3, so every line that's a multiple of 3
            // will display that line number to the left of the text
            for (int i = 1; i <= 25; i++)
                builder.Writeln($"Line {i}.");

            // The line counter will skip any paragraph with this flag set to true
            // Normally, the number "15" would normally appear next to this paragraph, which says "Line 15"
            // Since we set this flag to true and this paragraph is not counted by numbering,
            // number 15 will appear next to the next paragraph, "Line 16", and from then on counting will carry on as normal
            // until it will restart according to LineNumberRestartMode
            doc.FirstSection.Body.Paragraphs[14].ParagraphFormat.SuppressLineNumbers = true;

            doc.Save(ArtifactsDir + "PageSetup.LineNumbers.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.LineNumbers.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(1, pageSetup.LineStartingNumber);
            Assert.AreEqual(3, pageSetup.LineNumberCountBy);
            Assert.AreEqual(LineNumberRestartMode.RestartPage, pageSetup.LineNumberRestartMode);
            Assert.AreEqual(50.0d, pageSetup.LineNumberDistanceFromText);
        }

        [Test]
        public void PageBorderProperties()
        {
            //ExStart
            //ExFor:Section.PageSetup
            //ExFor:PageSetup.BorderAlwaysInFront
            //ExFor:PageSetup.BorderDistanceFrom
            //ExFor:PageSetup.BorderAppliesTo
            //ExFor:PageBorderDistanceFrom
            //ExFor:PageBorderAppliesTo
            //ExFor:Border.DistanceFromText
            //ExSummary:Shows how to create a page border that looks like a wide blue band at the top of the first page only.
            Document doc = new Document();

            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.BorderAlwaysInFront = false;
            pageSetup.BorderDistanceFrom = PageBorderDistanceFrom.PageEdge;
            pageSetup.BorderAppliesTo = PageBorderAppliesTo.FirstPage;

            Border border = pageSetup.Borders[BorderType.Top];
            border.LineStyle = LineStyle.Single;
            border.LineWidth = 30;
            border.Color = Color.Blue;
            border.DistanceFromText = 0;

            doc.Save(ArtifactsDir + "PageSetup.PageBorderProperties.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PageBorderProperties.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.False(pageSetup.BorderAlwaysInFront);
            Assert.AreEqual(PageBorderDistanceFrom.PageEdge, pageSetup.BorderDistanceFrom);
            Assert.AreEqual(PageBorderAppliesTo.FirstPage, pageSetup.BorderAppliesTo);

            border = pageSetup.Borders[BorderType.Top];

            Assert.AreEqual(LineStyle.Single, border.LineStyle);
            Assert.AreEqual(30.0d, border.LineWidth);
            Assert.AreEqual(Color.Blue.ToArgb(), border.Color.ToArgb());
            Assert.AreEqual(0.0d, border.DistanceFromText);
        }

        [Test]
        public void PageBorders()
        {
            //ExStart
            //ExFor:PageSetup.Borders
            //ExFor:Border.Shadow
            //ExFor:BorderCollection.LineStyle
            //ExFor:BorderCollection.LineWidth
            //ExFor:BorderCollection.Color
            //ExFor:BorderCollection.DistanceFromText
            //ExFor:BorderCollection.Shadow
            //ExSummary:Shows how to create green wavy page border with a shadow.
            Document doc = new Document();
            PageSetup pageSetup = doc.Sections[0].PageSetup;

            pageSetup.Borders.LineStyle = LineStyle.DoubleWave;
            pageSetup.Borders.LineWidth = 2;
            pageSetup.Borders.Color = Color.Green;
            pageSetup.Borders.DistanceFromText = 24;
            pageSetup.Borders.Shadow = true;

            doc.Save(ArtifactsDir + "PageSetup.PageBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PageBorders.docx");
            pageSetup = doc.FirstSection.PageSetup;

            foreach (Border border in pageSetup.Borders)
            {
                Assert.AreEqual(LineStyle.DoubleWave, border.LineStyle);
                Assert.AreEqual(2.0d, border.LineWidth);
                Assert.AreEqual(Color.Green.ToArgb(), border.Color.ToArgb());
                Assert.AreEqual(24.0d, border.DistanceFromText);
                Assert.True(border.Shadow);
            }
        }

        [Test]
        public void PageNumbering()
        {
            //ExStart
            //ExFor:PageSetup.RestartPageNumbering
            //ExFor:PageSetup.PageStartingNumber
            //ExFor:PageSetup.PageNumberStyle
            //ExFor:DocumentBuilder.InsertField(String, String)
            //ExSummary:Shows how to control page numbering per section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2");

            // Use document builder to create a header with a page number field for the first section
            // The page number will look like "Page V"
            builder.MoveToSection(0);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Page ");
            builder.InsertField("PAGE", "");

            // Set first section page numbering
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.RestartPageNumbering = true;
            pageSetup.PageStartingNumber = 5;
            pageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

            // Create a header for the section
            // The page number will look like " - 10 - ".
            builder.MoveToSection(1);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write(" - ");
            builder.InsertField("PAGE", "");
            builder.Write(" - ");

            // Set second section page numbering
            pageSetup = doc.Sections[1].PageSetup;
            pageSetup.PageStartingNumber = 10;
            pageSetup.RestartPageNumbering = true;
            pageSetup.PageNumberStyle = NumberStyle.Arabic;

            doc.Save(ArtifactsDir + "PageSetup.PageNumbering.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PageNumbering.docx");
            pageSetup = doc.Sections[0].PageSetup;

            Assert.True(pageSetup.RestartPageNumbering);
            Assert.AreEqual(5, pageSetup.PageStartingNumber);
            Assert.AreEqual(NumberStyle.UppercaseRoman, pageSetup.PageNumberStyle);

            pageSetup = doc.Sections[1].PageSetup;

            Assert.True(pageSetup.RestartPageNumbering);
            Assert.AreEqual(10, pageSetup.PageStartingNumber);
            Assert.AreEqual(NumberStyle.Arabic, pageSetup.PageNumberStyle);
        }

        [Test]
        public void FootnoteOptions()
        {
            //ExStart
            //ExFor:PageSetup.EndnoteOptions
            //ExFor:PageSetup.FootnoteOptions
            //ExSummary:Shows how to set options for footnotes and endnotes in current section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text and a reference for it in the form of a footnote
            builder.Write("Hello world!.");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote reference text.");

            // Set options for footnote position and numbering
            FootnoteOptions footnoteOptions = doc.Sections[0].PageSetup.FootnoteOptions;
            footnoteOptions.Position = FootnotePosition.BeneathText;
            footnoteOptions.RestartRule = FootnoteNumberingRule.RestartPage;
            footnoteOptions.StartNumber = 1;

            // Endnotes also have a similar options object
            builder.Write(" Hello again.");
            builder.InsertFootnote(FootnoteType.Footnote, "Endnote reference text.");

            EndnoteOptions endnoteOptions = doc.Sections[0].PageSetup.EndnoteOptions;
            endnoteOptions.Position = EndnotePosition.EndOfDocument;
            endnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;
            endnoteOptions.StartNumber = 1;

            doc.Save(ArtifactsDir + "PageSetup.FootnoteOptions.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.FootnoteOptions.docx");
            footnoteOptions = doc.FirstSection.PageSetup.FootnoteOptions;

            Assert.AreEqual(FootnotePosition.BeneathText, footnoteOptions.Position);
            Assert.AreEqual(FootnoteNumberingRule.RestartPage, footnoteOptions.RestartRule);
            Assert.AreEqual(1, footnoteOptions.StartNumber);

            endnoteOptions = doc.FirstSection.PageSetup.EndnoteOptions;

            Assert.AreEqual(EndnotePosition.EndOfDocument, endnoteOptions.Position);
            Assert.AreEqual(FootnoteNumberingRule.Continuous, endnoteOptions.RestartRule);
            Assert.AreEqual(1, endnoteOptions.StartNumber);
        }

        [Test]
        public void Bidi()
        {
            //ExStart
            //ExFor:PageSetup.Bidi
            //ExSummary:Shows how to change the order of columns.
            Document doc = new Document();

            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.TextColumns.SetCount(3);

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Write("Column 2.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Write("Column 3.");

            // Reverse the order of the columns
            pageSetup.Bidi = true;

            doc.Save(ArtifactsDir + "PageSetup.Bidi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.Bidi.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(3, pageSetup.TextColumns.Count);
            Assert.True(pageSetup.Bidi);
        }

        [Test]
        public void PageBorder()
        {
            //ExStart
            //ExFor:PageSetup.BorderSurroundsFooter
            //ExFor:PageSetup.BorderSurroundsHeader
            //ExSummary:Shows how to apply a border to the page and header/footer.
            Document doc = new Document();

            // Insert header and footer text
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Footer");
            builder.MoveToDocumentEnd();

            // Insert a page border and set the color and line style
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.Borders.LineStyle = LineStyle.Double;
            pageSetup.Borders.Color = Color.Blue;

            // By default, page borders don't surround headers and footers
            // We can change that by setting these flags
            pageSetup.BorderSurroundsFooter = true;
            pageSetup.BorderSurroundsHeader = true;

            doc.Save(ArtifactsDir + "PageSetup.PageBorder.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PageBorder.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.True(pageSetup.BorderSurroundsFooter);
            Assert.True(pageSetup.BorderSurroundsHeader);
        }

        [Test]
        public void Gutter()
        {
            //ExStart
            //ExFor:PageSetup.Gutter
            //ExFor:PageSetup.RtlGutter
            //ExFor:PageSetup.MultiplePages
            //ExSummary:Shows how to set gutter margins.
            Document doc = new Document();

            // Insert text spanning several pages
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (int i = 0; i < 6; i++)
            {
                builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
                builder.InsertBreak(BreakType.PageBreak);
            }

            // We can access the gutter margin in the section's page options,
            // which is a margin which is added to the page margin at one side of the page
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.Gutter = 100.0d;

            // If our text is LTR, the gutter will appear on the left side of the page
            // Setting this flag will move it to the right side
            pageSetup.RtlGutter = true;

            // Mirroring the margins will make the gutter alternate in position from page to page
            pageSetup.MultiplePages = MultiplePagesType.MirrorMargins;

            doc.Save(ArtifactsDir + "PageSetup.Gutter.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.Gutter.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(100.0d, pageSetup.Gutter);
            Assert.True(pageSetup.RtlGutter);
            Assert.AreEqual(MultiplePagesType.MirrorMargins, pageSetup.MultiplePages);
        }

        [Test]
        public void Booklet()
        {
            //ExStart
            //ExFor:PageSetup.SheetsPerBooklet
            //ExSummary:Shows how to create a booklet.
            Document doc = new Document();

            // Use a document builder to create 16 pages of content that will be compiled in a booklet
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("My Booklet:");

            for (int i = 0; i < 15; i++)
            {
                builder.InsertBreak(BreakType.PageBreak);
                builder.Write($"Booklet face #{i}");
            }

            // Set the number of sheets that will be used by the printer to create the booklet
            // After being printed on both sides, the sheets can be stacked and folded down the centre
            // The contents that we placed in such a way that they will be in order once the booklet is folded
            // We can only specify the number of sheets in multiples of 4
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;
            pageSetup.SheetsPerBooklet = 4;

            doc.Save(ArtifactsDir + "PageSetup.Booklet.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.Booklet.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(MultiplePagesType.BookFoldPrinting, pageSetup.MultiplePages);
            Assert.AreEqual(4, pageSetup.SheetsPerBooklet);
        }

        [Test]
        public void SectionTextOrientation()
        {
            //ExStart
            //ExFor:PageSetup.TextOrientation
            //ExSummary:Shows how to set text orientation.
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Setting this value will rotate the section's text 90 degrees to the right
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.TextOrientation = TextOrientation.Upward;

            doc.Save(ArtifactsDir + "PageSetup.SectionTextOrientation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.SectionTextOrientation.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.AreEqual(TextOrientation.Upward, pageSetup.TextOrientation);
        }

        //ExStart
        //ExFor:PageSetup.SuppressEndnotes
        //ExFor:Body.ParentSection
        //ExSummary:Shows how to store endnotes at the end of each section instead of the document and manipulate their positions.
        [Test] //ExSkip
        public void SuppressEndnotes()
        {
            // Create a new document and make it empty
            Document doc = new Document();
            doc.RemoveAllChildren();

            // Normally endnotes are all stored at the end of a document, but this option lets us store them at the end of each section
            doc.EndnoteOptions.Position = EndnotePosition.EndOfSection;

            // Create 3 new sections, each having a paragraph and an endnote at the end
            InsertSection(doc, "Section 1", "Endnote 1, will stay in section 1");
            InsertSection(doc, "Section 2", "Endnote 2, will be pushed down to section 3");
            InsertSection(doc, "Section 3", "Endnote 3, will stay in section 3");

            // Each section contains its own page setup object
            // Setting this value will push this section's endnotes down to the next section
            PageSetup pageSetup = doc.Sections[1].PageSetup;
            pageSetup.SuppressEndnotes = true;

            doc.Save(ArtifactsDir + "PageSetup.SuppressEndnotes.docx");
            TestSuppressEndnotes(new Document(ArtifactsDir + "PageSetup.SuppressEndnotes.docx")); //ExSkip
        }

        /// <summary>
        /// Add a section to the end of a document, give it a body and a paragraph, then add text and an endnote to that paragraph.
        /// </summary>
        private static void InsertSection(Document doc, string sectionBodyText, string endnoteText)
        {
            Section section = new Section(doc);

            doc.AppendChild(section);

            Body body = new Body(doc);
            section.AppendChild(body);

            Assert.AreEqual(section, body.ParentNode);

            Paragraph para = new Paragraph(doc);
            body.AppendChild(para);

            Assert.AreEqual(body, para.ParentNode);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(para);
            builder.Write(sectionBodyText);
            builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
        }
        //ExEnd

        private static void TestSuppressEndnotes(Document doc)
        {
            PageSetup pageSetup = doc.Sections[1].PageSetup;

            Assert.True(pageSetup.SuppressEndnotes);
        }
    }
}