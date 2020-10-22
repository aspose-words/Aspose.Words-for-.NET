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
            //ExSummary:Shows how to apply and revert page setup settings to sections in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Modify the page setup properties for the builder's current section, and add text.
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.VerticalAlignment = PageVerticalAlignment.Center;
            builder.Writeln("This is the first section, which landscape oriented with vertically centered text.");

            // If we start a new section using a document builder,
            // it will inherit the builder's current page setup properties.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            Assert.AreEqual(Orientation.Landscape, doc.Sections[1].PageSetup.Orientation);
            Assert.AreEqual(PageVerticalAlignment.Center, doc.Sections[1].PageSetup.VerticalAlignment);

            // We can revert its page setup properties to their default values using the "ClearFormatting" method.
            builder.PageSetup.ClearFormatting();

            Assert.AreEqual(Orientation.Portrait, doc.Sections[1].PageSetup.Orientation);
            Assert.AreEqual(PageVerticalAlignment.Top, doc.Sections[1].PageSetup.VerticalAlignment);

            builder.Writeln("This is the second section, which is in default Letter paper size, portrait orientation and top alignment.");

            doc.Save(ArtifactsDir + "PageSetup.ClearFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.ClearFormatting.docx");

            Assert.AreEqual(Orientation.Landscape, doc.Sections[0].PageSetup.Orientation);
            Assert.AreEqual(PageVerticalAlignment.Center, doc.Sections[0].PageSetup.VerticalAlignment);

            Assert.AreEqual(Orientation.Portrait, doc.Sections[1].PageSetup.Orientation);
            Assert.AreEqual(PageVerticalAlignment.Top, doc.Sections[1].PageSetup.VerticalAlignment);
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
            // 1 -  The "First" header/footer, which appear at the first page of the section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Writeln("First page header.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
            builder.Writeln("First page footer.");

            // 2 -  The "Primary" header/footer, which appears on every page in the section.
            // We can override the primary header/footer by a first, as well as an even page header/footer. 
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
            // such as orientation, size and borders.
            // Set the "DifferentFirstPageHeaderFooter" property to "true" to apply the first header/footer to the first page.
            // Set the "DifferentFirstPageHeaderFooter" property to "false"
            // to make the first page display the primary header/footer.
            builder.PageSetup.DifferentFirstPageHeaderFooter = differentFirstPageHeaderFooter;

            doc.Save(ArtifactsDir + "PageSetup.DifferentFirstPageHeaderFooter.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.DifferentFirstPageHeaderFooter.docx");

            Assert.AreEqual(differentFirstPageHeaderFooter, doc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter);
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
            // We can override the primary header/footer by a first, as well as an even page header/footer. 
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
            // such as orientation, size and borders.
            // Set the "OddAndEvenPagesHeaderFooter" property to "true"
            // to display the even page header/footer on even pages.
            // Set the "OddAndEvenPagesHeaderFooter" property to "false"
            // to display the primary header/footer on even pages.
            builder.PageSetup.OddAndEvenPagesHeaderFooter = oddAndEvenPagesHeaderFooter;

            doc.Save(ArtifactsDir + "PageSetup.OddAndEvenPagesHeaderFooter.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.OddAndEvenPagesHeaderFooter.docx");

            Assert.AreEqual(oddAndEvenPagesHeaderFooter, doc.FirstSection.PageSetup.OddAndEvenPagesHeaderFooter);
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

            Assert.AreEqual(8, doc.FirstSection.PageSetup.CharactersPerLine);

            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            doc.Save(ArtifactsDir + "PageSetup.CharactersPerLine.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.CharactersPerLine.docx");

            Assert.AreEqual(SectionLayoutMode.Grid, doc.FirstSection.PageSetup.LayoutMode);
            Assert.AreEqual(8, doc.FirstSection.PageSetup.CharactersPerLine);
        }

        [Test]
        public void LinesPerPage()
        {
            //ExStart
            //ExFor:PageSetup.LinesPerPage
            //ExFor:PageSetup.LayoutMode
            //ExFor:SectionLayoutMode
            //ExSummary:Shows how to specify a limit for the number of lines that each page may have.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable pitching, and then use it to set the number of lines per page in this section.
            // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
            builder.PageSetup.LayoutMode = SectionLayoutMode.LineGrid;
            builder.PageSetup.LinesPerPage = 15;

            for (int i = 0; i < 30; i++)
                builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");

            doc.Save(ArtifactsDir + "PageSetup.LinesPerPage.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.LinesPerPage.docx");

            Assert.AreEqual(SectionLayoutMode.LineGrid, doc.FirstSection.PageSetup.LayoutMode);
            Assert.AreEqual(15, doc.FirstSection.PageSetup.LinesPerPage);
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

            Assert.AreEqual(SectionStart.NewPage, doc.Sections[1].PageSetup.SectionStart);

            // 2 -  Starts the next section on the current page:
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.Writeln("This text is in section 3.");

            Assert.AreEqual(SectionStart.Continuous, doc.Sections[2].PageSetup.SectionStart);

            // 3 -  Starts the next section on a new even page:
            builder.InsertBreak(BreakType.SectionBreakEvenPage);
            builder.Writeln("This text is in section 4.");

            Assert.AreEqual(SectionStart.EvenPage, doc.Sections[3].PageSetup.SectionStart);

            // 4 -  Starts the next section on a new odd page:
            builder.InsertBreak(BreakType.SectionBreakOddPage);
            builder.Writeln("This text is in section 5.");

            Assert.AreEqual(SectionStart.OddPage, doc.Sections[4].PageSetup.SectionStart);

            // 5 -  Starts the next section on a new column:
            TextColumnCollection columns = builder.PageSetup.TextColumns;
            columns.SetCount(2);

            builder.InsertBreak(BreakType.SectionBreakNewColumn);
            builder.Writeln("This text is in section 6.");

            Assert.AreEqual(SectionStart.NewColumn, doc.Sections[5].PageSetup.SectionStart);

            doc.Save(ArtifactsDir + "PageSetup.SetSectionStart.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.SetSectionStart.docx");

            Assert.AreEqual(SectionStart.NewPage, doc.Sections[0].PageSetup.SectionStart);
            Assert.AreEqual(SectionStart.NewPage, doc.Sections[1].PageSetup.SectionStart);
            Assert.AreEqual(SectionStart.Continuous, doc.Sections[2].PageSetup.SectionStart);
            Assert.AreEqual(SectionStart.EvenPage, doc.Sections[3].PageSetup.SectionStart);
            Assert.AreEqual(SectionStart.OddPage, doc.Sections[4].PageSetup.SectionStart);
            Assert.AreEqual(SectionStart.NewColumn, doc.Sections[5].PageSetup.SectionStart);
        }

#if NET462 || NETCOREAPP2_1 || JAVA
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
            
            // The paper tray value stored in documents is printer-specific.
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

            // Find the default printer that we will use for printing this document.
            // You can define a specific printer using the "PrinterName" property of the PrinterSettings object.
            PrinterSettings settings = new PrinterSettings();

            // This is the tray we will use for pages in the "A4" paper size.
            int printerTrayForA4 = settings.PaperSources[0].RawKind;

            // The is the tray we will use for pages in the "Letter" paper size.
            int printerTrayForLetter = settings.PaperSources[1].RawKind;

            // Modify the PageSettings object of this section to get Microsoft Word to instruct the printer
            // to use one the trays we identified above, depending on the paper size of this section.
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

            Assert.AreEqual(PaperSize.Legal, doc.FirstSection.PageSetup.PaperSize);
            Assert.AreEqual(1008.0d, doc.FirstSection.PageSetup.PageWidth);
            Assert.AreEqual(612.0d, doc.FirstSection.PageSetup.PageHeight);
            Assert.AreEqual(Orientation.Landscape, doc.FirstSection.PageSetup.Orientation);
            Assert.AreEqual(72.0d, doc.FirstSection.PageSetup.TopMargin);
            Assert.AreEqual(72.0d, doc.FirstSection.PageSetup.BottomMargin);
            Assert.AreEqual(108.0d, doc.FirstSection.PageSetup.LeftMargin);
            Assert.AreEqual(108.0d, doc.FirstSection.PageSetup.RightMargin);
            Assert.AreEqual(14.4d, doc.FirstSection.PageSetup.HeaderDistance);
            Assert.AreEqual(14.4d, doc.FirstSection.PageSetup.FooterDistance);
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

            // We can change the size of the current page to a pre-defined size
            // by using the "PaperSize" property of this section's PageSetup object.
            builder.PageSetup.PaperSize = PaperSize.Tabloid;

            Assert.AreEqual(792.0d, builder.PageSetup.PageWidth);
            Assert.AreEqual(1224.0d, builder.PageSetup.PageHeight);

            builder.Writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

            // Each section has its own PageSetup object. When we use a document builder to make a new section,
            // that section's PageSetup object inherits all the values of the previous section's PageSetup object.
            builder.InsertBreak(BreakType.SectionBreakEvenPage);

            Assert.AreEqual(PaperSize.Tabloid, builder.PageSetup.PaperSize);

            builder.PageSetup.PaperSize = PaperSize.A5;
            builder.Writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

            Assert.AreEqual(419.55d, builder.PageSetup.PageWidth);
            Assert.AreEqual(595.30d, builder.PageSetup.PageHeight);

            builder.InsertBreak(BreakType.SectionBreakEvenPage);

            // Set a custom size for this section's pages.
            builder.PageSetup.PageWidth = 620;
            builder.PageSetup.PageHeight = 480;

            Assert.AreEqual(PaperSize.Custom, builder.PageSetup.PaperSize);

            builder.Writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

            doc.Save(ArtifactsDir + "PageSetup.PaperSizes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "PageSetup.PaperSizes.docx");

            Assert.AreEqual(PaperSize.Tabloid, doc.Sections[0].PageSetup.PaperSize);
            Assert.AreEqual(792.0d, doc.Sections[0].PageSetup.PageWidth);
            Assert.AreEqual(1224.0d, doc.Sections[0].PageSetup.PageHeight);
            Assert.AreEqual(PaperSize.A5, doc.Sections[1].PageSetup.PaperSize);
            Assert.AreEqual(419.55d, doc.Sections[1].PageSetup.PageWidth);
            Assert.AreEqual(595.30d, doc.Sections[1].PageSetup.PageHeight);
            Assert.AreEqual(PaperSize.Custom, doc.Sections[2].PageSetup.PaperSize);
            Assert.AreEqual(620.0d, doc.Sections[2].PageSetup.PageWidth);
            Assert.AreEqual(480.0d, doc.Sections[2].PageSetup.PageHeight);
        }

        [Test]
        public void ColumnsSameWidth()
        {
            //ExStart
            //ExFor:PageSetup.TextColumns
            //ExFor:TextColumnCollection
            //ExFor:TextColumnCollection.Spacing
            //ExFor:TextColumnCollection.SetCount
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

            Assert.AreEqual(100.0d, doc.FirstSection.PageSetup.TextColumns.Spacing);
            Assert.AreEqual(2, doc.FirstSection.PageSetup.TextColumns.Count);
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

            Assert.AreEqual(470.30d, contentWidth, 0.01d);

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

            Assert.False(pageSetup.TextColumns.EvenlySpaced);
            Assert.AreEqual(2, pageSetup.TextColumns.Count);
            Assert.AreEqual(100.0d, pageSetup.TextColumns[0].Width);
            Assert.AreEqual(20.0d, pageSetup.TextColumns[0].SpaceAfter);
            Assert.AreEqual(470.3d, pageSetup.TextColumns[1].Width);
            Assert.AreEqual(0.0d, pageSetup.TextColumns[1].SpaceAfter);
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

            Assert.AreEqual(lineBetween, doc.FirstSection.PageSetup.TextColumns.LineBetween);
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

            // We can use the section's PageSetup object to display numbers to the left of lines of text in the section.
            // This is the same behavior as a List object,
            // but it covers the entire section, and does not modify the text in any way.
            // Our section will restart the numbering on each new page from 1, and will display the number,
            // if it is a multiple of 3, at a distance of 50pt to the left of the line.
            PageSetup pageSetup = builder.PageSetup;
            pageSetup.LineStartingNumber = 1;
            pageSetup.LineNumberCountBy = 3;
            pageSetup.LineNumberRestartMode = LineNumberRestartMode.RestartPage;
            pageSetup.LineNumberDistanceFromText = 50.0d;

            for (int i = 1; i <= 25; i++)
                builder.Writeln($"Line {i}.");

            // The line counter will skip any paragraph with the "SuppressLineNumbers" flag set to "true".
            // This paragraph is on the 15th line, which is a multiple of 3, and thus would normally display a line number.
            // The section's line counter will also ignore this line, and treat the next line as the 15th,
            // and continue the count from that point onward.
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

            // By default, page borders do not surround headers and footers
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
            // After being printed on both sides, the sheets can be stacked and folded down the middle
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