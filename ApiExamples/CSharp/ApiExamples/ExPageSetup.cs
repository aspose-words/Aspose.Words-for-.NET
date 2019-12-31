// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Settings;
using NUnit.Framework;
#if NETFRAMEWORK || NETSTANDARD2_0
using System.Drawing.Printing;
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
            DocumentBuilder builder = new DocumentBuilder();

            // Modify the first section in the document
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.VerticalAlignment = PageVerticalAlignment.Center;
            builder.Writeln("Section 1, landscape oriented and text vertically centered.");

            // Start a new section and reset its formatting to defaults
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.PageSetup.ClearFormatting();
            builder.Writeln("Section 2, back to default Letter paper size, portrait orientation and top alignment.");

            builder.Document.Save(ArtifactsDir + "PageSetup.ClearFormatting.doc");
            //ExEnd
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
            //ExSummary:Creates headers and footers different for first, even and odd pages using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder();

            PageSetup ps = builder.PageSetup;
            ps.DifferentFirstPageHeaderFooter = true;
            ps.OddAndEvenPagesHeaderFooter = true;
            ps.LayoutMode = SectionLayoutMode.LineGrid;
            ps.CharactersPerLine = 1;
            ps.LinesPerPage = 1;

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

            builder.Document.Save(ArtifactsDir + "PageSetup.DifferentHeaders.doc");
            //ExEnd
        }

        [Test]
        public void SectionStart()
        {
            //ExStart
            //ExFor:SectionStart
            //ExFor:PageSetup.SectionStart
            //ExFor:Document.Sections
            //ExSummary:Specifies how the section starts, from a new page, on the same page or other.
            Document doc = new Document();
            doc.Sections[0].PageSetup.SectionStart = Aspose.Words.SectionStart.Continuous;
            //ExEnd
        }

#if NETFRAMEWORK || NETSTANDARD2_0
        [Test]
        [Ignore("Run only when the printer driver is installed")]
        public void DefaultPaperTray()
        {
            //ExStart
            //ExFor:PageSetup.FirstPageTray
            //ExFor:PageSetup.OtherPagesTray
            //ExSummary:Changes all sections in a document to use the default paper tray of the selected printer.
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
            //ExSummary:Specifies paper size, orientation, margins and other settings for a section.
            DocumentBuilder builder = new DocumentBuilder();

            PageSetup ps = builder.PageSetup;
            ps.PaperSize = Aspose.Words.PaperSize.Legal;
            ps.Orientation = Orientation.Landscape;
            ps.TopMargin = ConvertUtil.InchToPoint(1.0);
            ps.BottomMargin = ConvertUtil.InchToPoint(1.0);
            ps.LeftMargin = ConvertUtil.InchToPoint(1.5);
            ps.RightMargin = ConvertUtil.InchToPoint(1.5);
            ps.HeaderDistance = ConvertUtil.InchToPoint(0.2);
            ps.FooterDistance = ConvertUtil.InchToPoint(0.2);

            builder.Writeln("Hello world.");

            builder.Document.Save(ArtifactsDir + "PageSetup.PageMargins.doc");
            //ExEnd
        }

        [Test]
        public void ColumnsSameWidth()
        {
            //ExStart
            //ExFor:PageSetup.TextColumns
            //ExFor:TextColumnCollection
            //ExFor:TextColumnCollection.Spacing
            //ExFor:TextColumnCollection.SetCount
            //ExSummary:Creates multiple evenly spaced columns in a section using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder();

            TextColumnCollection columns = builder.PageSetup.TextColumns;
            // Make spacing between columns wider
            columns.Spacing = 100;
            // This creates two columns of equal width
            columns.SetCount(2);

            builder.Writeln("Text in column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Text in column 2.");

            builder.Document.Save(ArtifactsDir + "PageSetup.ColumnsSameWidth.doc");
            //ExEnd
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
            //ExSummary:Creates multiple columns of different widths in a section using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder();

            TextColumnCollection columns = builder.PageSetup.TextColumns;
            // Show vertical line between columns
            columns.LineBetween = true;
            // Indicate we want to create column with different widths
            columns.EvenlySpaced = false;
            // Create two columns, note they will be created with zero widths, need to set them
            columns.SetCount(2);

            // Set the first column to be narrow
            TextColumn c1 = columns[0];
            c1.Width = 100;
            c1.SpaceAfter = 20;

            // Set the second column to take the rest of the space available on the page
            TextColumn c2 = columns[1];
            PageSetup ps = builder.PageSetup;
            double contentWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin;
            c2.Width = contentWidth - c1.Width - c1.SpaceAfter;

            builder.Writeln("Narrow column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Wide column 2.");

            builder.Document.Save(ArtifactsDir + "PageSetup.CustomColumnWidth.doc");
            //ExEnd
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
            //ExSummary:Turns on Microsoft Word line numbering for a section.
            DocumentBuilder builder = new DocumentBuilder();

            PageSetup ps = builder.PageSetup;
            ps.LineStartingNumber = 1;
            ps.LineNumberCountBy = 5;
            ps.LineNumberRestartMode = LineNumberRestartMode.RestartPage;
            ps.LineNumberDistanceFromText = 50.0d;

            // The line counter will skip any paragraph with this flag set to true
            Assert.False(builder.ParagraphFormat.SuppressLineNumbers);

            for (int i = 1; i <= 20; i++)
                builder.Writeln($"Line {i}.");

            builder.Document.Save(ArtifactsDir + "PageSetup.LineNumbers.docx");
            //ExEnd
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
            //ExSummary:Creates a page border that looks like a wide blue band at the top of the first page only.
            Document doc = new Document();

            PageSetup ps = doc.Sections[0].PageSetup;
            ps.BorderAlwaysInFront = false;
            ps.BorderDistanceFrom = PageBorderDistanceFrom.PageEdge;
            ps.BorderAppliesTo = PageBorderAppliesTo.FirstPage;

            Border border = ps.Borders[BorderType.Top];
            border.LineStyle = LineStyle.Single;
            border.LineWidth = 30;
            border.Color = Color.Blue;
            border.DistanceFromText = 0;

            doc.Save(ArtifactsDir + "PageSetup.PageBorderProperties.doc");
            //ExEnd
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
            //ExSummary:Creates a fancy looking green wavy page border with a shadow.
            Document doc = new Document();
            PageSetup ps = doc.Sections[0].PageSetup;

            ps.Borders.LineStyle = LineStyle.DoubleWave;
            ps.Borders.LineWidth = 2;
            ps.Borders.Color = Color.Green;
            ps.Borders.DistanceFromText = 24;
            ps.Borders.Shadow = true;

            doc.Save(ArtifactsDir + "PageSetup.PageBorders.doc");
            //ExEnd
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
            // This document has two sections, but no page numbers yet
            Document doc = new Document(MyDir + "PageSetup.PageNumbering.doc");

            // Use document builder to create a header with a page number field for the first section
            // The page number will look like "Page V"
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToSection(0);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Page ");
            builder.InsertField("PAGE", "");

            // Set first section page numbering
            Section section = doc.Sections[0];
            section.PageSetup.RestartPageNumbering = true;
            section.PageSetup.PageStartingNumber = 5;
            section.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

            // Create a header for the section
            // The page number will look like " - 10 - ".
            builder.MoveToSection(1);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write(" - ");
            builder.InsertField("PAGE", "");
            builder.Write(" - ");

            // Set second section page numbering
            section = doc.Sections[1];
            section.PageSetup.PageStartingNumber = 10;
            section.PageSetup.RestartPageNumbering = true;
            section.PageSetup.PageNumberStyle = NumberStyle.Arabic;

            doc.Save(ArtifactsDir + "PageSetup.PageNumbering.doc");
            //ExEnd
        }

        [Test]
        public void FootnoteOptions()
        {
            //ExStart
            //ExFor:PageSetup.FootnoteOptions
            //ExSummary:Shows how to set options for footnotes in current section.
            Document doc = new Document();

            PageSetup pageSetup = doc.Sections[0].PageSetup;

            pageSetup.FootnoteOptions.Position = FootnotePosition.BottomOfPage;
            pageSetup.FootnoteOptions.NumberStyle = NumberStyle.Bullet;
            pageSetup.FootnoteOptions.StartNumber = 1;
            pageSetup.FootnoteOptions.RestartRule = FootnoteNumberingRule.RestartPage;
            pageSetup.FootnoteOptions.Columns = 0;
            //ExEnd
        }

        [Test]
        public void EndnoteOptions()
        {
            //ExStart
            //ExFor:PageSetup.EndnoteOptions
            //ExSummary:Shows how to set options for endnotes in current section.
            Document doc = new Document();

            PageSetup pageSetup = doc.Sections[0].PageSetup;

            pageSetup.EndnoteOptions.Position = EndnotePosition.EndOfSection;
            pageSetup.EndnoteOptions.NumberStyle = NumberStyle.Bullet;
            pageSetup.EndnoteOptions.StartNumber = 1;
            pageSetup.EndnoteOptions.RestartRule = FootnoteNumberingRule.RestartPage;
            //ExEnd
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
        }

        [Test]
        public void BorderSurrounds()
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

            doc.Save(ArtifactsDir + "PageSetup.BorderSurrounds.docx");
            //ExEnd
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
            pageSetup.Gutter = 100.0;

            // If our text is LTR, the gutter will appear on the left side of the page
            // Setting this flag will move it to the right side
            pageSetup.RtlGutter = true;

            // Mirroring the margins will make the gutter alternate in position from page to page
            pageSetup.MultiplePages = MultiplePagesType.MirrorMargins;

            doc.Save(ArtifactsDir + "PageSetup.Gutter.docx");
            //ExEnd
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
        }

        [Test]
        public void TextOrientation()
        {
            //ExStart
            //ExFor:PageSetup.TextOrientation
            //ExSummary:Shows how to set text orientation.
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Setting this value will rotate the section's text 90 degrees to the right
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.TextOrientation = Aspose.Words.TextOrientation.Upward;

            doc.Save(ArtifactsDir + "PageSetup.TextOrientation.docx");
            //ExEnd
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
    }
}