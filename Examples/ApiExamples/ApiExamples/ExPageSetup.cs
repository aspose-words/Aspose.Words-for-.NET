// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Rendering;
using Aspose.Words.Settings;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using PaperSize = Aspose.Words.PaperSize;

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
            //ExSummary:Shows how to apply and revert page setup settings to sections in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Modify the page setup properties for the builder's current section and add text.
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.VerticalAlignment = PageVerticalAlignment.Center;
            builder.Writeln("This is the first section, which landscape oriented with vertically centered text.");

            // If we start a new section using a document builder,
            // it will inherit the builder's current page setup properties.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            Assert.That(doc.Sections[1].PageSetup.Orientation, Is.EqualTo(Orientation.Landscape));
            Assert.That(doc.Sections[1].PageSetup.VerticalAlignment, Is.EqualTo(PageVerticalAlignment.Center));

            // We can revert its page setup properties to their default values using the "ClearFormatting" method.
            builder.PageSetup.ClearFormatting();

            Assert.That(doc.Sections[1].PageSetup.Orientation, Is.EqualTo(Orientation.Portrait));
            Assert.That(doc.Sections[1].PageSetup.VerticalAlignment, Is.EqualTo(PageVerticalAlignment.Top));

            builder.Writeln("This is the second section, which is in default Letter paper size, portrait orientation and top alignment.");

            doc.Save(ArtifactsDir + "PageSetup.ClearFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.ClearFormatting.docx");

            Assert.That(doc.Sections[0].PageSetup.Orientation, Is.EqualTo(Orientation.Landscape));
            Assert.That(doc.Sections[0].PageSetup.VerticalAlignment, Is.EqualTo(PageVerticalAlignment.Center));

            Assert.That(doc.Sections[1].PageSetup.Orientation, Is.EqualTo(Orientation.Portrait));
            Assert.That(doc.Sections[1].PageSetup.VerticalAlignment, Is.EqualTo(PageVerticalAlignment.Top));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void DifferentFirstPageHeaderFooter(bool differentFirstPageHeaderFooter)
        {
            //ExStart
            //ExFor:PageSetup.DifferentFirstPageHeaderFooter
            //ExSummary:Shows how to enable or disable primary headers/footers.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two types of header/footers.
            // 1 -  The "First" header/footer, which appears on the first page of the section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Writeln("First page header.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
            builder.Writeln("First page footer.");

            // 2 -  The "Primary" header/footer, which appears on every page in the section.
            // We can override the primary header/footer by a first and an even page header/footer. 
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Primary header.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("Primary footer.");

            builder.MoveToSection(0);
            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            // Each section has a "PageSetup" object that specifies page appearance-related properties
            // such as orientation, size, and borders.
            // Set the "DifferentFirstPageHeaderFooter" property to "true" to apply the first header/footer to the first page.
            // Set the "DifferentFirstPageHeaderFooter" property to "false"
            // to make the first page display the primary header/footer.
            builder.PageSetup.DifferentFirstPageHeaderFooter = differentFirstPageHeaderFooter;

            doc.Save(ArtifactsDir + "PageSetup.DifferentFirstPageHeaderFooter.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.DifferentFirstPageHeaderFooter.docx");

            Assert.That(doc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter, Is.EqualTo(differentFirstPageHeaderFooter));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void OddAndEvenPagesHeaderFooter(bool oddAndEvenPagesHeaderFooter)
        {
            //ExStart
            //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
            //ExSummary:Shows how to enable or disable even page headers/footers.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two types of header/footers.
            // 1 -  The "Primary" header/footer, which appears on every page in the section.
            // We can override the primary header/footer by a first and an even page header/footer. 
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Primary header.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("Primary footer.");

            // 2 -  The "Even" header/footer, which appears on every even page of this section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Writeln("Even page header.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
            builder.Writeln("Even page footer.");

            builder.MoveToSection(0);
            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            // Each section has a "PageSetup" object that specifies page appearance-related properties
            // such as orientation, size, and borders.
            // Set the "OddAndEvenPagesHeaderFooter" property to "true"
            // to display the even page header/footer on even pages.
            // Set the "OddAndEvenPagesHeaderFooter" property to "false"
            // to display the primary header/footer on even pages.
            builder.PageSetup.OddAndEvenPagesHeaderFooter = oddAndEvenPagesHeaderFooter;

            doc.Save(ArtifactsDir + "PageSetup.OddAndEvenPagesHeaderFooter.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.OddAndEvenPagesHeaderFooter.docx");

            Assert.That(doc.FirstSection.PageSetup.OddAndEvenPagesHeaderFooter, Is.EqualTo(oddAndEvenPagesHeaderFooter));
        }

        [Test]
        public void CharactersPerLine()
        {
            //ExStart
            //ExFor:PageSetup.CharactersPerLine
            //ExFor:PageSetup.LayoutMode
            //ExFor:SectionLayoutMode
            //ExSummary:Shows how to specify a for the number of characters that each line may have.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable pitching, and then use it to set the number of characters per line in this section.
            builder.PageSetup.LayoutMode = SectionLayoutMode.Grid;
            builder.PageSetup.CharactersPerLine = 10;

            // The number of characters also depends on the size of the font.
            doc.Styles["Normal"].Font.Size = 20;

            Assert.That(doc.FirstSection.PageSetup.CharactersPerLine, Is.EqualTo(8));

            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            doc.Save(ArtifactsDir + "PageSetup.CharactersPerLine.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.CharactersPerLine.docx");

            Assert.That(doc.FirstSection.PageSetup.LayoutMode, Is.EqualTo(SectionLayoutMode.Grid));
            Assert.That(doc.FirstSection.PageSetup.CharactersPerLine, Is.EqualTo(8));
        }

        [Test]
        public void LinesPerPage()
        {
            //ExStart
            //ExFor:PageSetup.LinesPerPage
            //ExFor:PageSetup.LayoutMode
            //ExFor:ParagraphFormat.SnapToGrid
            //ExFor:SectionLayoutMode
            //ExSummary:Shows how to specify a limit for the number of lines that each page may have.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable pitching, and then use it to set the number of lines per page in this section.
            // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
            builder.PageSetup.LayoutMode = SectionLayoutMode.LineGrid;
            builder.PageSetup.LinesPerPage = 15;

            builder.ParagraphFormat.SnapToGrid = true;

            for (int i = 0; i < 30; i++)
                builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");

            doc.Save(ArtifactsDir + "PageSetup.LinesPerPage.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.LinesPerPage.docx");

            Assert.That(doc.FirstSection.PageSetup.LayoutMode, Is.EqualTo(SectionLayoutMode.LineGrid));
            Assert.That(doc.FirstSection.PageSetup.LinesPerPage, Is.EqualTo(15));

            foreach (Paragraph paragraph in doc.FirstSection.Body.Paragraphs)
                Assert.That(paragraph.ParagraphFormat.SnapToGrid, Is.True);
        }

        [Test]
        public void SetSectionStart()
        {
            //ExStart
            //ExFor:SectionStart
            //ExFor:PageSetup.SectionStart
            //ExFor:Document.Sections
            //ExSummary:Shows how to specify how a new section separates itself from the previous.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This text is in section 1.");

            // Section break types determine how a new section separates itself from the previous section.
            // Below are five types of section breaks.
            // 1 -  Starts the next section on a new page:
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("This text is in section 2.");

            Assert.That(doc.Sections[1].PageSetup.SectionStart, Is.EqualTo(SectionStart.NewPage));

            // 2 -  Starts the next section on the current page:
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.Writeln("This text is in section 3.");

            Assert.That(doc.Sections[2].PageSetup.SectionStart, Is.EqualTo(SectionStart.Continuous));

            // 3 -  Starts the next section on a new even page:
            builder.InsertBreak(BreakType.SectionBreakEvenPage);
            builder.Writeln("This text is in section 4.");

            Assert.That(doc.Sections[3].PageSetup.SectionStart, Is.EqualTo(SectionStart.EvenPage));

            // 4 -  Starts the next section on a new odd page:
            builder.InsertBreak(BreakType.SectionBreakOddPage);
            builder.Writeln("This text is in section 5.");

            Assert.That(doc.Sections[4].PageSetup.SectionStart, Is.EqualTo(SectionStart.OddPage));

            // 5 -  Starts the next section on a new column:
            TextColumnCollection columns = builder.PageSetup.TextColumns;
            columns.SetCount(2);

            builder.InsertBreak(BreakType.SectionBreakNewColumn);
            builder.Writeln("This text is in section 6.");

            Assert.That(doc.Sections[5].PageSetup.SectionStart, Is.EqualTo(SectionStart.NewColumn));

            doc.Save(ArtifactsDir + "PageSetup.SetSectionStart.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.SetSectionStart.docx");

            Assert.That(doc.Sections[0].PageSetup.SectionStart, Is.EqualTo(SectionStart.NewPage));
            Assert.That(doc.Sections[1].PageSetup.SectionStart, Is.EqualTo(SectionStart.NewPage));
            Assert.That(doc.Sections[2].PageSetup.SectionStart, Is.EqualTo(SectionStart.Continuous));
            Assert.That(doc.Sections[3].PageSetup.SectionStart, Is.EqualTo(SectionStart.EvenPage));
            Assert.That(doc.Sections[4].PageSetup.SectionStart, Is.EqualTo(SectionStart.OddPage));
            Assert.That(doc.Sections[5].PageSetup.SectionStart, Is.EqualTo(SectionStart.NewColumn));
        }

        [Test]
        [Ignore("Run only when the printer driver is installed")]
        public void DefaultPaperTray()
        {
            //ExStart
            //ExFor:PageSetup.FirstPageTray
            //ExFor:PageSetup.OtherPagesTray
            //ExSummary:Shows how to get all the sections in a document to use the default paper tray of the selected printer.
            Document doc = new Document();

            // Find the default printer that we will use for printing this document.
            // You can define a specific printer using the "PrinterName" property of the PrinterSettings object.
            PrinterSettings settings = new PrinterSettings();

            // The paper tray value stored in documents is printer specific.
            // This means the code below resets all page tray values to use the current printers default tray.
            // You can enumerate PrinterSettings.PaperSources to find the other valid paper tray values of the selected printer.
            foreach (Section section in doc.Sections.OfType<Section>())
            {
                section.PageSetup.FirstPageTray = settings.DefaultPageSettings.PaperSource.RawKind;
                section.PageSetup.OtherPagesTray = settings.DefaultPageSettings.PaperSource.RawKind;
            }
            //ExEnd

            foreach (Section section in DocumentHelper.SaveOpen(doc).Sections.OfType<Section>())
            {
                Assert.That(section.PageSetup.FirstPageTray, Is.EqualTo(settings.DefaultPageSettings.PaperSource.RawKind));
                Assert.That(section.PageSetup.OtherPagesTray, Is.EqualTo(settings.DefaultPageSettings.PaperSource.RawKind));
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

            // Find the default printer that we will use for printing this document.
            // You can define a specific printer using the "PrinterName" property of the PrinterSettings object.
            PrinterSettings settings = new PrinterSettings();

            // This is the tray we will use for pages in the "A4" paper size.
            int printerTrayForA4 = settings.PaperSources[0].RawKind;

            // This is the tray we will use for pages in the "Letter" paper size.
            int printerTrayForLetter = settings.PaperSources[1].RawKind;

            // Modify the PageSettings object of this section to get Microsoft Word to instruct the printer
            // to use one of the trays we identified above, depending on this section's paper size.
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
                    Assert.That(section.PageSetup.FirstPageTray, Is.EqualTo(printerTrayForLetter));
                    Assert.That(section.PageSetup.OtherPagesTray, Is.EqualTo(printerTrayForLetter));
                }
                else if (section.PageSetup.PaperSize == Aspose.Words.PaperSize.A4)
                {
                    Assert.That(section.PageSetup.FirstPageTray, Is.EqualTo(printerTrayForA4));
                    Assert.That(section.PageSetup.OtherPagesTray, Is.EqualTo(printerTrayForA4));
                }
            }
        }

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
            //ExSummary:Shows how to adjust paper size, orientation, margins, along with other settings for a section.
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

            builder.Writeln("Hello world!");

            doc.Save(ArtifactsDir + "PageSetup.PageMargins.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PageMargins.docx");

            Assert.That(doc.FirstSection.PageSetup.PaperSize, Is.EqualTo(PaperSize.Legal));
            Assert.That(doc.FirstSection.PageSetup.PageWidth, Is.EqualTo(1008.0d));
            Assert.That(doc.FirstSection.PageSetup.PageHeight, Is.EqualTo(612.0d));
            Assert.That(doc.FirstSection.PageSetup.Orientation, Is.EqualTo(Orientation.Landscape));
            Assert.That(doc.FirstSection.PageSetup.TopMargin, Is.EqualTo(72.0d));
            Assert.That(doc.FirstSection.PageSetup.BottomMargin, Is.EqualTo(72.0d));
            Assert.That(doc.FirstSection.PageSetup.LeftMargin, Is.EqualTo(108.0d));
            Assert.That(doc.FirstSection.PageSetup.RightMargin, Is.EqualTo(108.0d));
            Assert.That(doc.FirstSection.PageSetup.HeaderDistance, Is.EqualTo(14.4d));
            Assert.That(doc.FirstSection.PageSetup.FooterDistance, Is.EqualTo(14.4d));
        }

        [Test]
        public void PaperSizes()
        {
            //ExStart
            //ExFor:PaperSize
            //ExFor:PageSetup.PaperSize
            //ExSummary:Shows how to set page sizes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We can change the current page's size to a pre-defined size
            // by using the "PaperSize" property of this section's PageSetup object.
            builder.PageSetup.PaperSize = PaperSize.Tabloid;

            Assert.That(builder.PageSetup.PageWidth, Is.EqualTo(792.0d));
            Assert.That(builder.PageSetup.PageHeight, Is.EqualTo(1224.0d));

            builder.Writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

            // Each section has its own PageSetup object. When we use a document builder to make a new section,
            // that section's PageSetup object inherits all the previous section's PageSetup object's values.
            builder.InsertBreak(BreakType.SectionBreakEvenPage);

            Assert.That(builder.PageSetup.PaperSize, Is.EqualTo(PaperSize.Tabloid));

            builder.PageSetup.PaperSize = PaperSize.A5;
            builder.Writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

            Assert.That(builder.PageSetup.PageWidth, Is.EqualTo(419.55d));
            Assert.That(builder.PageSetup.PageHeight, Is.EqualTo(595.30d));

            builder.InsertBreak(BreakType.SectionBreakEvenPage);

            // Set a custom size for this section's pages.
            builder.PageSetup.PageWidth = 620;
            builder.PageSetup.PageHeight = 480;

            Assert.That(builder.PageSetup.PaperSize, Is.EqualTo(PaperSize.Custom));

            builder.Writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

            doc.Save(ArtifactsDir + "PageSetup.PaperSizes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PaperSizes.docx");

            Assert.That(doc.Sections[0].PageSetup.PaperSize, Is.EqualTo(PaperSize.Tabloid));
            Assert.That(doc.Sections[0].PageSetup.PageWidth, Is.EqualTo(792.0d));
            Assert.That(doc.Sections[0].PageSetup.PageHeight, Is.EqualTo(1224.0d));
            Assert.That(doc.Sections[1].PageSetup.PaperSize, Is.EqualTo(PaperSize.A5));
            Assert.That(doc.Sections[1].PageSetup.PageWidth, Is.EqualTo(419.55d));
            Assert.That(doc.Sections[1].PageSetup.PageHeight, Is.EqualTo(595.30d));
            Assert.That(doc.Sections[2].PageSetup.PaperSize, Is.EqualTo(PaperSize.Custom));
            Assert.That(doc.Sections[2].PageSetup.PageWidth, Is.EqualTo(620.0d));
            Assert.That(doc.Sections[2].PageSetup.PageHeight, Is.EqualTo(480.0d));
        }

        [Test]
        public void ColumnsSameWidth()
        {
            //ExStart
            //ExFor:PageSetup.TextColumns
            //ExFor:TextColumnCollection
            //ExFor:TextColumnCollection.Spacing
            //ExFor:TextColumnCollection.SetCount
            //ExFor:TextColumnCollection.Count
            //ExFor:TextColumnCollection.Width
            //ExSummary:Shows how to create multiple evenly spaced columns in a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            TextColumnCollection columns = builder.PageSetup.TextColumns;
            columns.Spacing = 100;
            columns.SetCount(2);

            builder.Writeln("Column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Column 2.");

            doc.Save(ArtifactsDir + "PageSetup.ColumnsSameWidth.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.ColumnsSameWidth.docx");

            Assert.That(doc.FirstSection.PageSetup.TextColumns.Spacing, Is.EqualTo(100.0d));
            Assert.That(doc.FirstSection.PageSetup.TextColumns.Count, Is.EqualTo(2));
            Assert.That(doc.FirstSection.PageSetup.TextColumns.Width, Is.EqualTo(185.15).Within(0.01));
        }

        [Test]
        public void CustomColumnWidth()
        {
            //ExStart
            //ExFor:TextColumnCollection.EvenlySpaced
            //ExFor:TextColumnCollection.Item
            //ExFor:TextColumn
            //ExFor:TextColumn.Width
            //ExFor:TextColumn.SpaceAfter
            //ExSummary:Shows how to create unevenly spaced columns.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            PageSetup pageSetup = builder.PageSetup;

            TextColumnCollection columns = pageSetup.TextColumns;
            columns.EvenlySpaced = false;
            columns.SetCount(2);

            // Determine the amount of room that we have available for arranging columns.
            double contentWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;

            Assert.That(contentWidth, Is.EqualTo(470.30d).Within(0.01d));

            // Set the first column to be narrow.
            TextColumn column = columns[0];
            column.Width = 100;
            column.SpaceAfter = 20;

            // Set the second column to take the rest of the space available within the margins of the page.
            column = columns[1];
            column.Width = contentWidth - column.Width - column.SpaceAfter;

            builder.Writeln("Narrow column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Wide column 2.");

            doc.Save(ArtifactsDir + "PageSetup.CustomColumnWidth.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.CustomColumnWidth.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.TextColumns.EvenlySpaced, Is.False);
            Assert.That(pageSetup.TextColumns.Count, Is.EqualTo(2));
            Assert.That(pageSetup.TextColumns[0].Width, Is.EqualTo(100.0d));
            Assert.That(pageSetup.TextColumns[0].SpaceAfter, Is.EqualTo(20.0d));
            Assert.That(pageSetup.TextColumns[1].Width, Is.EqualTo(470.3d));
            Assert.That(pageSetup.TextColumns[1].SpaceAfter, Is.EqualTo(0.0d));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void VerticalLineBetweenColumns(bool lineBetween)
        {
            //ExStart
            //ExFor:TextColumnCollection.LineBetween
            //ExSummary:Shows how to separate columns with a vertical line.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Configure the current section's PageSetup object to divide the text into several columns.
            // Set the "LineBetween" property to "true" to put a dividing line between columns.
            // Set the "LineBetween" property to "false" to leave the space between columns blank.
            TextColumnCollection columns = builder.PageSetup.TextColumns;
            columns.LineBetween = lineBetween;
            columns.SetCount(3);

            builder.Writeln("Column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Column 2.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Column 3.");

            doc.Save(ArtifactsDir + "PageSetup.VerticalLineBetweenColumns.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.VerticalLineBetweenColumns.docx");

            Assert.That(doc.FirstSection.PageSetup.TextColumns.LineBetween, Is.EqualTo(lineBetween));
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
            //ExSummary:Shows how to enable line numbering for a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We can use the section's PageSetup object to display numbers to the left of the section's text lines.
            // This is the same behavior as a List object,
            // but it covers the entire section and does not modify the text in any way.
            // Our section will restart the numbering on each new page from 1 and display the number,
            // if it is a multiple of 3, at 50pt to the left of the line.
            PageSetup pageSetup = builder.PageSetup;
            pageSetup.LineStartingNumber = 1;
            pageSetup.LineNumberCountBy = 3;
            pageSetup.LineNumberRestartMode = LineNumberRestartMode.RestartPage;
            pageSetup.LineNumberDistanceFromText = 50.0d;

            for (int i = 1; i <= 25; i++)
                builder.Writeln($"Line {i}.");

            // The line counter will skip any paragraph with the "SuppressLineNumbers" flag set to "true".
            // This paragraph is on the 15th line, which is a multiple of 3, and thus would normally display a line number.
            // The section's line counter will also ignore this line, treat the next line as the 15th,
            // and continue the count from that point onward.
            doc.FirstSection.Body.Paragraphs[14].ParagraphFormat.SuppressLineNumbers = true;

            doc.Save(ArtifactsDir + "PageSetup.LineNumbers.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.LineNumbers.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.LineStartingNumber, Is.EqualTo(1));
            Assert.That(pageSetup.LineNumberCountBy, Is.EqualTo(3));
            Assert.That(pageSetup.LineNumberRestartMode, Is.EqualTo(LineNumberRestartMode.RestartPage));
            Assert.That(pageSetup.LineNumberDistanceFromText, Is.EqualTo(50.0d));
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
            //ExSummary:Shows how to create a wide blue band border at the top of the first page.
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

            Assert.That(pageSetup.BorderAlwaysInFront, Is.False);
            Assert.That(pageSetup.BorderDistanceFrom, Is.EqualTo(PageBorderDistanceFrom.PageEdge));
            Assert.That(pageSetup.BorderAppliesTo, Is.EqualTo(PageBorderAppliesTo.FirstPage));

            border = pageSetup.Borders[BorderType.Top];

            Assert.That(border.LineStyle, Is.EqualTo(LineStyle.Single));
            Assert.That(border.LineWidth, Is.EqualTo(30.0d));
            Assert.That(border.Color.ToArgb(), Is.EqualTo(Color.Blue.ToArgb()));
            Assert.That(border.DistanceFromText, Is.EqualTo(0.0d));
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
                Assert.That(border.LineStyle, Is.EqualTo(LineStyle.DoubleWave));
                Assert.That(border.LineWidth, Is.EqualTo(2.0d));
                Assert.That(border.Color.ToArgb(), Is.EqualTo(Color.Green.ToArgb()));
                Assert.That(border.DistanceFromText, Is.EqualTo(24.0d));
                Assert.That(border.Shadow, Is.True);
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
            //ExSummary:Shows how to set up page numbering in a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Section 1, page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Section 1, page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Section 1, page 3.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2, page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Section 2, page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Section 2, page 3.");

            // Move the document builder to the first section's primary header,
            // which every page in that section will display.
            builder.MoveToSection(0);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a PAGE field, which will display the number of the current page.
            builder.Write("Page ");
            builder.InsertField("PAGE", "");

            // Configure the section to have the page count that PAGE fields display start from 5.
            // Also, configure all PAGE fields to display their page numbers using uppercase Roman numerals.
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.RestartPageNumbering = true;
            pageSetup.PageStartingNumber = 5;
            pageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;

            // Create another primary header for the second section, with another PAGE field.
            builder.MoveToSection(1);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write(" - ");
            builder.InsertField("PAGE", "");
            builder.Write(" - ");

            // Configure the section to have the page count that PAGE fields display start from 10.
            // Also, configure all PAGE fields to display their page numbers using Arabic numbers.
            pageSetup = doc.Sections[1].PageSetup;
            pageSetup.PageStartingNumber = 10;
            pageSetup.RestartPageNumbering = true;
            pageSetup.PageNumberStyle = NumberStyle.Arabic;

            doc.Save(ArtifactsDir + "PageSetup.PageNumbering.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PageNumbering.docx");
            pageSetup = doc.Sections[0].PageSetup;

            Assert.That(pageSetup.RestartPageNumbering, Is.True);
            Assert.That(pageSetup.PageStartingNumber, Is.EqualTo(5));
            Assert.That(pageSetup.PageNumberStyle, Is.EqualTo(NumberStyle.UppercaseRoman));

            pageSetup = doc.Sections[1].PageSetup;

            Assert.That(pageSetup.RestartPageNumbering, Is.True);
            Assert.That(pageSetup.PageStartingNumber, Is.EqualTo(10));
            Assert.That(pageSetup.PageNumberStyle, Is.EqualTo(NumberStyle.Arabic));
        }

        [Test]
        public void FootnoteOptions()
        {
            //ExStart
            //ExFor:PageSetup.EndnoteOptions
            //ExFor:PageSetup.FootnoteOptions
            //ExSummary:Shows how to configure options affecting footnotes/endnotes in a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote reference text.");

            // Configure all footnotes in the first section to restart the numbering from 1
            // at each new page and display themselves directly beneath the text on every page.
            FootnoteOptions footnoteOptions = doc.Sections[0].PageSetup.FootnoteOptions;
            footnoteOptions.Position = FootnotePosition.BeneathText;
            footnoteOptions.RestartRule = FootnoteNumberingRule.RestartPage;
            footnoteOptions.StartNumber = 1;

            builder.Write(" Hello again.");
            builder.InsertFootnote(FootnoteType.Footnote, "Endnote reference text.");

            // Configure all endnotes in the first section to maintain a continuous count throughout the section,
            // starting from 1. Also, set them all to appear collected at the end of the document.
            EndnoteOptions endnoteOptions = doc.Sections[0].PageSetup.EndnoteOptions;
            endnoteOptions.Position = EndnotePosition.EndOfDocument;
            endnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;
            endnoteOptions.StartNumber = 1;

            doc.Save(ArtifactsDir + "PageSetup.FootnoteOptions.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.FootnoteOptions.docx");
            footnoteOptions = doc.FirstSection.PageSetup.FootnoteOptions;

            Assert.That(footnoteOptions.Position, Is.EqualTo(FootnotePosition.BeneathText));
            Assert.That(footnoteOptions.RestartRule, Is.EqualTo(FootnoteNumberingRule.RestartPage));
            Assert.That(footnoteOptions.StartNumber, Is.EqualTo(1));

            endnoteOptions = doc.FirstSection.PageSetup.EndnoteOptions;

            Assert.That(endnoteOptions.Position, Is.EqualTo(EndnotePosition.EndOfDocument));
            Assert.That(endnoteOptions.RestartRule, Is.EqualTo(FootnoteNumberingRule.Continuous));
            Assert.That(endnoteOptions.StartNumber, Is.EqualTo(1));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void Bidi(bool reverseColumns)
        {
            //ExStart
            //ExFor:PageSetup.Bidi
            //ExSummary:Shows how to set the order of text columns in a section.
            Document doc = new Document();

            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.TextColumns.SetCount(3);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Write("Column 2.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Write("Column 3.");

            // Set the "Bidi" property to "true" to arrange the columns starting from the page's right side.
            // The order of the columns will match the direction of the right-to-left text.
            // Set the "Bidi" property to "false" to arrange the columns starting from the page's left side.
            // The order of the columns will match the direction of the left-to-right text.
            pageSetup.Bidi = reverseColumns;

            doc.Save(ArtifactsDir + "PageSetup.Bidi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.Bidi.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.TextColumns.Count, Is.EqualTo(3));
            Assert.That(pageSetup.Bidi, Is.EqualTo(reverseColumns));
        }

        [Test]
        public void PageBorder()
        {
            //ExStart
            //ExFor:PageSetup.BorderSurroundsFooter
            //ExFor:PageSetup.BorderSurroundsHeader
            //ExSummary:Shows how to apply a border to the page and header/footer.
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! This is the main body text.");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("This is the header.");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("This is the footer.");
            builder.MoveToDocumentEnd();

            // Insert a blue double-line border.
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.Borders.LineStyle = LineStyle.Double;
            pageSetup.Borders.Color = Color.Blue;

            // A section's PageSetup object has "BorderSurroundsHeader" and "BorderSurroundsFooter" flags that determine
            // whether a page border surrounds the main body text, also includes the header or footer, respectively.
            // Set the "BorderSurroundsHeader" flag to "true" to surround the header with our border,
            // and then set the "BorderSurroundsFooter" flag to leave the footer outside of the border.
            pageSetup.BorderSurroundsHeader = true;
            pageSetup.BorderSurroundsFooter = false;

            doc.Save(ArtifactsDir + "PageSetup.PageBorder.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PageBorder.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.BorderSurroundsHeader, Is.True);
            Assert.That(pageSetup.BorderSurroundsFooter, Is.False);
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

            // Insert text that spans several pages.
            DocumentBuilder builder = new DocumentBuilder(doc);
            for (int i = 0; i < 6; i++)
            {
                builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                              "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
                builder.InsertBreak(BreakType.PageBreak);
            }

            // A gutter adds whitespaces to either the left or right page margin,
            // which makes up for the center folding of pages in a book encroaching on the page's layout.
            PageSetup pageSetup = doc.Sections[0].PageSetup;

            // Determine how much space our pages have for text within the margins and then add an amount to pad a margin. 
            Assert.That(pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin, Is.EqualTo(470.30d).Within(0.01d));

            pageSetup.Gutter = 100.0d;

            // Set the "RtlGutter" property to "true" to place the gutter in a more suitable position for right-to-left text.
            pageSetup.RtlGutter = true;

            // Set the "MultiplePages" property to "MultiplePagesType.MirrorMargins" to alternate
            // the left/right page side position of margins every page.
            pageSetup.MultiplePages = MultiplePagesType.MirrorMargins;

            doc.Save(ArtifactsDir + "PageSetup.Gutter.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.Gutter.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.Gutter, Is.EqualTo(100.0d));
            Assert.That(pageSetup.RtlGutter, Is.True);
            Assert.That(pageSetup.MultiplePages, Is.EqualTo(MultiplePagesType.MirrorMargins));
        }

        [Test]
        public void Booklet()
        {
            //ExStart
            //ExFor:PageSetup.Gutter
            //ExFor:PageSetup.MultiplePages
            //ExFor:PageSetup.SheetsPerBooklet
            //ExFor:MultiplePagesType
            //ExSummary:Shows how to configure a document that can be printed as a book fold.
            Document doc = new Document();

            // Insert text that spans 16 pages.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("My Booklet:");

            for (int i = 0; i < 15; i++)
            {
                builder.InsertBreak(BreakType.PageBreak);
                builder.Write($"Booklet face #{i}");
            }

            // Configure the first section's "PageSetup" property to print the document in the form of a book fold.
            // When we print this document on both sides, we can take the pages to stack them
            // and fold them all down the middle at once. The contents of the document will line up into a book fold.
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;

            // We can only specify the number of sheets in multiples of 4.
            pageSetup.SheetsPerBooklet = 4;

            doc.Save(ArtifactsDir + "PageSetup.Booklet.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.Booklet.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.MultiplePages, Is.EqualTo(MultiplePagesType.BookFoldPrinting));
            Assert.That(pageSetup.SheetsPerBooklet, Is.EqualTo(4));
        }

        [Test]
        public void SetTextOrientation()
        {
            //ExStart
            //ExFor:PageSetup.TextOrientation
            //ExSummary:Shows how to set text orientation.
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Set the "TextOrientation" property to "TextOrientation.Upward" to rotate all the text 90 degrees
            // to the right so that all left-to-right text now goes top-to-bottom.
            PageSetup pageSetup = doc.Sections[0].PageSetup;
            pageSetup.TextOrientation = TextOrientation.Upward;

            doc.Save(ArtifactsDir + "PageSetup.SetTextOrientation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.SetTextOrientation.docx");
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.TextOrientation, Is.EqualTo(TextOrientation.Upward));
        }

        //ExStart
        //ExFor:PageSetup.SuppressEndnotes
        //ExFor:Body.ParentSection
        //ExSummary:Shows how to store endnotes at the end of each section, and modify their positions.
        [Test] //ExSkip
        public void SuppressEndnotes()
        {
            Document doc = new Document();
            doc.RemoveAllChildren();

            // By default, a document compiles all endnotes at its end. 
            Assert.That(doc.EndnoteOptions.Position, Is.EqualTo(EndnotePosition.EndOfDocument));

            // We use the "Position" property of the document's "EndnoteOptions" object
            // to collect endnotes at the end of each section instead. 
            doc.EndnoteOptions.Position = EndnotePosition.EndOfSection;

            InsertSectionWithEndnote(doc, "Section 1", "Endnote 1, will stay in section 1");
            InsertSectionWithEndnote(doc, "Section 2", "Endnote 2, will be pushed down to section 3");
            InsertSectionWithEndnote(doc, "Section 3", "Endnote 3, will stay in section 3");

            // While getting sections to display their respective endnotes, we can set the "SuppressEndnotes" flag
            // of a section's "PageSetup" object to "true" to revert to the default behavior and pass its endnotes
            // onto the next section.
            PageSetup pageSetup = doc.Sections[1].PageSetup;
            pageSetup.SuppressEndnotes = true;

            doc.Save(ArtifactsDir + "PageSetup.SuppressEndnotes.docx");
            TestSuppressEndnotes(new Document(ArtifactsDir + "PageSetup.SuppressEndnotes.docx")); //ExSkip
        }

        /// <summary>
        /// Append a section with text and an endnote to a document.
        /// </summary>
        private static void InsertSectionWithEndnote(Document doc, string sectionBodyText, string endnoteText)
        {
            Section section = new Section(doc);

            doc.AppendChild(section);

            Body body = new Body(doc);
            section.AppendChild(body);

            Assert.That(body.ParentNode, Is.EqualTo(section));

            Paragraph para = new Paragraph(doc);
            body.AppendChild(para);

            Assert.That(para.ParentNode, Is.EqualTo(body));

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(para);
            builder.Write(sectionBodyText);
            builder.InsertFootnote(FootnoteType.Endnote, endnoteText);
        }
        //ExEnd

        private static void TestSuppressEndnotes(Document doc)
        {
            PageSetup pageSetup = doc.Sections[1].PageSetup;

            Assert.That(pageSetup.SuppressEndnotes, Is.True);
        }

        [Test]
        public void ChapterPageSeparator()
        {
            //ExStart
            //ExFor:PageSetup.HeadingLevelForChapter
            //ExFor:ChapterPageSeparator
            //ExFor:PageSetup.ChapterPageSeparator
            //ExSummary:Shows how to work with page chapters.
            Document doc = new Document(MyDir + "Big document.docx");

            PageSetup pageSetup = doc.FirstSection.PageSetup;

            pageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;
            pageSetup.ChapterPageSeparator = Aspose.Words.ChapterPageSeparator.Colon;
            pageSetup.HeadingLevelForChapter = 1;
            //ExEnd
        }

        [Test]
        public void JisbPaperSize()
        {
            //ExStart:JisbPaperSize
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:PageSetup.PaperSize
            //ExSummary:Shows how to set the paper size of JisB4 or JisB5.
            Document doc = new Document(MyDir + "Big document.docx");

            PageSetup pageSetup = doc.FirstSection.PageSetup;
            // Set the paper size to JisB4 (257x364mm).
            pageSetup.PaperSize = PaperSize.JisB4;
            // Alternatively, set the paper size to JisB5. (182x257mm).
            pageSetup.PaperSize = PaperSize.JisB5;
            //ExEnd:JisbPaperSize

            doc = DocumentHelper.SaveOpen(doc);
            pageSetup = doc.FirstSection.PageSetup;

            Assert.That(pageSetup.PaperSize, Is.EqualTo(PaperSize.JisB5));
        }

#if NETFRAMEWORK
        [Test]
        [Ignore("Run only when the printer driver is installed")]
        public void PrintPagesRemaining()
        {
            //ExStart:PrintPagesRemaining
            //GistId:571cc6e23284a2ec075d15d4c32e3bbf
            //ExFor:AsposeWordsPrintDocument
            //ExFor:AsposeWordsPrintDocument.PagesRemaining
            //ExSummary: Shows how to monitor printing progress.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Initialize the printer settings.
            PrinterSettings printerSettings = new PrinterSettings();
            printerSettings.PrinterName = "Microsoft Print to PDF";
            printerSettings.PrintRange = PrintRange.AllPages;

            // Create a special Aspose.Words implementation of the .NET PrintDocument class.
            // Pass the printer settings from the print dialog to the print document.
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
            printDoc.PrinterSettings = printerSettings;

            // Initialize the custom printing tracker.
            PrintTracker printTracker = new PrintTracker(printDoc);

            printDoc.Print();

            // Write the event log.
            foreach (string eventString in printTracker.EventLog)
                Console.WriteLine(eventString);
            //ExEnd:PrintPagesRemaining
        }
#endif
    }
}