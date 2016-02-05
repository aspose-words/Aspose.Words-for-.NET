// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;


namespace ApiExamples.PageSetup
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

            // Modify the first section in the document.
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.VerticalAlignment = PageVerticalAlignment.Center;
            builder.Writeln("Section 1, landscape oriented and text vertically centered.");

            // Start a new section and reset its formatting to defaults.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.PageSetup.ClearFormatting();
            builder.Writeln("Section 2, back to default Letter paper size, portrait orientation and top alignment.");

            builder.Document.Save(MyDir + "PageSetup.ClearFormatting Out.doc");
            //ExEnd
        }

        [Test]
        public void DifferentHeaders()
        {
            //ExStart
            //ExFor:PageSetup.DifferentFirstPageHeaderFooter
            //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
            //ExSummary:Creates headers and footers different for first, even and odd pages using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder();

            Aspose.Words.PageSetup ps = builder.PageSetup;
            ps.DifferentFirstPageHeaderFooter = true;
            ps.OddAndEvenPagesHeaderFooter = true;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Writeln("First page header.");

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Writeln("Even pages header.");

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Odd pages header.");

            // Move back to the main story of the first section.
            builder.MoveToSection(0);
            builder.Writeln("Text page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Text page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Text page 3.");

            builder.Document.Save(MyDir + "PageSetup.DifferentHeaders Out.doc");
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
            Aspose.Words.Document doc = new Aspose.Words.Document();
            doc.Sections[0].PageSetup.SectionStart = Aspose.Words.SectionStart.Continuous;
            //ExEnd
        }

        [Test]
        public void DefaultPaperTray()
        {
            //ExStart
            //ExFor:PageSetup.FirstPageTray
            //ExFor:PageSetup.OtherPagesTray
            //ExSummary:Changes all sections in a document to use the default paper tray of the selected printer.
            Aspose.Words.Document doc = new Aspose.Words.Document();

            // Find the printer that will be used for printing this document. In this case it is the default printer.
            // You can define a specific printer using PrinterName.
            System.Drawing.Printing.PrinterSettings settings = new System.Drawing.Printing.PrinterSettings();

            // The paper tray value stored in documents is completely printer specific. This means 
            // The code below resets all page tray values to use the current printers default tray.
            // You can enumerate PrinterSettings.PaperSources to find the other valid paper tray values of the selected printer.
            foreach (Aspose.Words.Section section in doc.Sections)
            {
                section.PageSetup.FirstPageTray = settings.DefaultPageSettings.PaperSource.RawKind;
                section.PageSetup.OtherPagesTray = settings.DefaultPageSettings.PaperSource.RawKind;
            }
            //ExEnd
        }

        [Test]
        public void PaperTrayForDifferentPaperType()
        {
            //ExStart
            //ExFor:PageSetup.FirstPageTray
            //ExFor:PageSetup.OtherPagesTray
            //ExSummary:Shows how to set up printing using different printer trays for different paper sizes.
            Aspose.Words.Document doc = new Aspose.Words.Document();

            // Choose the default printer to be used for printing this document.
            System.Drawing.Printing.PrinterSettings settings = new System.Drawing.Printing.PrinterSettings();

            // This is the tray we will use for A4 paper size. This is the first tray in the paper sources collection.
            int printerTrayForA4 = settings.PaperSources[0].RawKind;
            // The is the tray we will use for Letter paper size. This is the second tray in the paper sources collection.
            int printerTrayForLetter = settings.PaperSources[1].RawKind;

            // Set the page tray used for each section based off the paper size used in the section.
            foreach (Aspose.Words.Section section in doc.Sections)
            {
                if(section.PageSetup.PaperSize == PaperSize.Letter)
                {
                    section.PageSetup.FirstPageTray = printerTrayForLetter;
                    section.PageSetup.OtherPagesTray = printerTrayForLetter;
                }
                else if(section.PageSetup.PaperSize == PaperSize.A4)
                {
                    section.PageSetup.FirstPageTray = printerTrayForA4;
                    section.PageSetup.OtherPagesTray = printerTrayForA4;
                }
            }
            //ExEnd
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
            //ExSummary:Specifies paper size, orientation, margins and other settings for a section.
            DocumentBuilder builder = new DocumentBuilder();

            Aspose.Words.PageSetup ps = builder.PageSetup;
            ps.PaperSize = PaperSize.Legal;
            ps.Orientation = Orientation.Landscape;
            ps.TopMargin = Aspose.Words.ConvertUtil.InchToPoint(1.0);
            ps.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(1.0);
            ps.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(1.5);
            ps.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(1.5);
            ps.HeaderDistance = Aspose.Words.ConvertUtil.InchToPoint(0.2);
            ps.FooterDistance = Aspose.Words.ConvertUtil.InchToPoint(0.2);

            builder.Writeln("Hello world.");

            builder.Document.Save(MyDir + "PageSetup.PageMargins Out.doc");
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
            // Make spacing between columns wider.
            columns.Spacing = 100;
            // This creates two columns of equal width.
            columns.SetCount(2);

            builder.Writeln("Text in column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Text in column 2.");

            builder.Document.Save(MyDir + "PageSetup.ColumnsSameWidth Out.doc");
            //ExEnd
        }

        [Test]
        public void ColumnsCustomWidth()
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
            // Show vertical line between columns.
            columns.LineBetween = true;
            // Indicate we want to create column with different widths.
            columns.EvenlySpaced = false;
            // Create two columns, note they will be created with zero widths, need to set them.
            columns.SetCount(2);

            // Set the first column to be narrow.
            TextColumn c1 = columns[0];
            c1.Width = 100;
            c1.SpaceAfter = 20;

            // Set the second column to take the rest of the space available on the page.
            TextColumn c2 = columns[1];
            Aspose.Words.PageSetup ps = builder.PageSetup;
            double contentWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin;
            c2.Width = contentWidth - c1.Width - c1.SpaceAfter;

            builder.Writeln("Narrow column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Wide column 2.");

            builder.Document.Save(MyDir + "PageSetup.ColumnsCustomWidth Out.doc");
            //ExEnd
        }

        [Test]
        public void LineNumbers()
        {
            //ExStart
            //ExFor:PageSetup.LineStartingNumber
            //ExFor:PageSetup.LineNumberCountBy
            //ExFor:PageSetup.LineNumberRestartMode
            //ExFor:LineNumberRestartMode
            //ExSummary:Turns on Microsoft Word line numbering for a section.
            DocumentBuilder builder = new DocumentBuilder();

            Aspose.Words.PageSetup ps = builder.PageSetup;
            ps.LineStartingNumber = 1;
            ps.LineNumberCountBy = 5;
            ps.LineNumberRestartMode = LineNumberRestartMode.RestartPage;

            for (int i = 1; i <= 20; i++)
                builder.Writeln(string.Format("Line {0}.", i));

            builder.Document.Save(MyDir + "PageSetup.LineNumbers Out.doc");
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
            Aspose.Words.Document doc = new Aspose.Words.Document();

            Aspose.Words.PageSetup ps = doc.Sections[0].PageSetup;
            ps.BorderAlwaysInFront = false;
            ps.BorderDistanceFrom = PageBorderDistanceFrom.PageEdge;
            ps.BorderAppliesTo = PageBorderAppliesTo.FirstPage;

            Aspose.Words.Border border = ps.Borders[BorderType.Top];
            border.LineStyle = LineStyle.Single;
            border.LineWidth = 30;
            border.Color = System.Drawing.Color.Blue;
            border.DistanceFromText = 0;

            doc.Save(MyDir + "PageSetup.PageBorderTop Out.doc");
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
            Aspose.Words.Document doc = new Aspose.Words.Document();
            Aspose.Words.PageSetup ps = doc.Sections[0].PageSetup;

            ps.Borders.LineStyle = LineStyle.DoubleWave;
            ps.Borders.LineWidth = 2;
            ps.Borders.Color = System.Drawing.Color.Green;
            ps.Borders.DistanceFromText = 24;
            ps.Borders.Shadow = true;

            doc.Save(MyDir + "PageSetup.PageBorders Out.doc");
            //ExEnd
        }

        [Test]
        public void PageNumbering()
        {
            //ExStart
            //ExFor:PageSetup.RestartPageNumbering
            //ExFor:PageSetup.PageStartingNumber
            //ExFor:PageSetup.PageNumberStyle
            //ExFor:DocumentBuilder.InsertField(string, string)
            //ExSummary:Shows how to control page numbering per section.
            // This document has two sections, but no page numbers yet.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "PageSetup.PageNumbering.doc");

            // Use document builder to create a header with a page number field for the first section.
            // The page number will look like "Page V".
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToSection(0);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Page ");
            builder.InsertField("PAGE", "");

            // Set first section page numbering.
            Aspose.Words.Section section = doc.Sections[0];
            section.PageSetup.RestartPageNumbering = true;
            section.PageSetup.PageStartingNumber = 5;
            section.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;


            // Create a header for the section section. 
            // The page number will look like " - 10 - ".
            builder.MoveToSection(1);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write(" - ");
            builder.InsertField("PAGE", "");
            builder.Write(" - ");

            // Set second section page numbering.
            section = doc.Sections[1];
            section.PageSetup.PageStartingNumber = 10;
            section.PageSetup.RestartPageNumbering = true;
            section.PageSetup.PageNumberStyle = NumberStyle.Arabic;

            doc.Save(MyDir + "PageSetup.PageNumbering Out.doc");
            //ExEnd
        }
    }
}
