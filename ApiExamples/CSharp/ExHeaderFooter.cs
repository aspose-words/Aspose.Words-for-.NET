// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExHeaderFooter : ApiExampleBase
    {
        [Test]
        public void CreateFooter()
        {
            //ExStart
            //ExFor:HeaderFooter
            //ExFor:HeaderFooter.#ctor(DocumentBase, HeaderFooterType)
            //ExFor:HeaderFooterCollection
            //ExFor:Story.AppendParagraph
            //ExSummary:Creates a footer using the document object model and inserts it into a section.
            Document doc = new Document();

            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            doc.FirstSection.HeadersFooters.Add(footer);

            // Add a paragraph with text to the footer.
            footer.AppendParagraph("TEST FOOTER");

            doc.Save(MyDir + @"\Artifacts\HeaderFooter.CreateFooter.doc");
            //ExEnd

            doc = new Document(MyDir + @"\Artifacts\HeaderFooter.CreateFooter.doc");
            Assert.True(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].Range.Text.Contains("TEST FOOTER"));
        }

        [Test]
        public void RemoveFooters()
        {
            //ExStart
            //ExFor:Section.HeadersFooters
            //ExFor:HeaderFooterCollection
            //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
            //ExFor:HeaderFooter
            //ExFor:HeaderFooterType
            //ExId:RemoveFooters
            //ExSummary:Deletes all footers from all sections, but leaves headers intact.
            Document doc = new Document(MyDir + "HeaderFooter.RemoveFooters.doc");

            foreach (Section section in doc)
            {
                // Up to three different footers are possible in a section (for first, even and odd pages).
                // We check and delete all of them.
                HeaderFooter footer;

                footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
                if (footer != null)
                    footer.Remove();

                // Primary footer is the footer used for odd pages.
                footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                    footer.Remove();

                footer = section.HeadersFooters[HeaderFooterType.FooterEven];
                if (footer != null)
                    footer.Remove();
            }

            doc.Save(MyDir + @"\Artifacts\HeaderFooter.RemoveFooters.doc");
            //ExEnd
        }

        [Test]
        public void SetExportHeadersFootersMode()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportHeadersFootersMode
            //ExFor:ExportHeadersFootersMode
            //ExSummary:Demonstrates how to disable the export of headers and footers when saving to HTML based formats.
            Document doc = new Document(MyDir + "HeaderFooter.RemoveFooters.doc");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.None; // Disables exporting headers and footers.

            doc.Save(MyDir + @"\Artifacts\HeaderFooter.DisableHeadersFooters.html", saveOptions);
            //ExEnd

            // Verify that the output document is correct.
            doc = new Document(MyDir + @"\Artifacts\HeaderFooter.DisableHeadersFooters.html");
            Assert.IsFalse(doc.Range.Text.Contains("DYNAMIC TEMPLATE"));
        }

        [Test]
        public void ReplaceText()
        {
            //ExStart
            //ExFor:Document.FirstSection
            //ExFor:Section.HeadersFooters
            //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
            //ExFor:HeaderFooter
            //ExFor:Range.Replace(String, String, Boolean, Boolean)
            //ExSummary:Shows how to replace text in the document footer.
            // Open the template document, containing obsolete copyright information in the footer.
            Document doc = new Document(MyDir + "HeaderFooter.ReplaceText.doc");

            HeaderFooterCollection headersFooters = doc.FirstSection.HeadersFooters;
            HeaderFooter footer = headersFooters[HeaderFooterType.FooterPrimary];
            footer.Range.Replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2011 by Aspose Pty Ltd.", false, false);

            doc.Save(MyDir + @"\Artifacts\HeaderFooter.ReplaceText.doc");
            //ExEnd

            // Verify that the appropriate changes were made to the output document.
            doc = new Document(MyDir + @"\Artifacts\HeaderFooter.ReplaceText.doc");
            Assert.IsTrue(doc.Range.Text.Contains("Copyright (C) 2011 by Aspose Pty Ltd."));
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void HeaderFooterPrimerCaller()
        {
            this.Primer();
        }

        //ExStart
        //ExId:HeaderFooterPrimer
        //ExSummary:Maybe a bit complicated example, but demonstrates many things that can be done with headers/footers.
        public void Primer()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;

            // Specify if we want headers/footers of the first page to be different from other pages.
            // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
            // different headers/footers for odd and even pages.
            pageSetup.DifferentFirstPageHeaderFooter = true;

            // --- Create header for the first page. ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Set font properties for header text.
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            // Specify header title for the first page.
            builder.Write("Aspose.Words Header/Footer Creation Primer - Title Page.");

            // --- Create header for pages other than first. ---
            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert absolutely positioned image into the top/left corner of the header.
            // Distance from the top/left edges of the page is set to 10 points.
            string imageFileName = MyDir + "Aspose.Words.gif";
            builder.InsertImage(imageFileName, RelativeHorizontalPosition.Page, 10, RelativeVerticalPosition.Page, 10, 50, 50, WrapType.Through);

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            // Specify another header title for other pages.
            builder.Write("Aspose.Words Header/Footer Creation Primer.");

            // --- Create footer for pages other than first. ---
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // We use table with two cells to make one part of the text on the line (with page numbering)
            // to be aligned left, and the other part of the text (with copyright) to be aligned right.
            builder.StartTable();

            // Clear table borders.
            builder.CellFormat.ClearFormatting();

            builder.InsertCell();

            // Set first cell to 1/3 of the page width.
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);

            // Insert page numbering text here.
            // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages.
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            // Align this text to the left.
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            builder.InsertCell();
            // Set the second cell to 2/3 of the page width.
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            builder.Write("(C) 2001 Aspose Pty Ltd. All rights reserved.");

            // Align this text to the right.
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            builder.EndRow();
            builder.EndTable();

            builder.MoveToDocumentEnd();
            // Make page break to create a second page on which the primary headers/footers will be seen.
            builder.InsertBreak(BreakType.PageBreak);

            // Make section break to create a third page with different page orientation.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Get the new section and its page setup.
            currentSection = builder.CurrentSection;
            pageSetup = currentSection.PageSetup;

            // Set page orientation of the new section to landscape.
            pageSetup.Orientation = Orientation.Landscape;

            // This section does not need different first page header/footer.
            // We need only one title page in the document and the header/footer for this page
            // has already been defined in the previous section
            pageSetup.DifferentFirstPageHeaderFooter = false;

            // This section displays headers/footers from the previous section by default.
            // Call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this.
            // Page width is different for the new section and therefore we need to set 
            // a different cell widths for a footer table.
            currentSection.HeadersFooters.LinkToPrevious(false);

            // If we want to use the already existing header/footer set for this section 
            // but with some minor modifications then it may be expedient to copy headers/footers
            // from the previous section and apply the necessary modifications where we want them.
            CopyHeadersFootersFromPreviousSection(currentSection);

            // Find the footer that we want to change.
            HeaderFooter primaryFooter = currentSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            Row row = primaryFooter.Tables[0].FirstRow;
            row.FirstCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);
            row.LastCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            // Save the resulting document.
            doc.Save(MyDir + @"\Artifacts\HeaderFooter.Primer.doc");
        }

        /// <summary>
        /// Clones and copies headers/footers form the previous section to the specified section.
        /// </summary>
        private static void CopyHeadersFootersFromPreviousSection(Section section)
        {
            Section previousSection = (Section)section.PreviousSibling;

            if (previousSection == null)
                return;

            section.HeadersFooters.Clear();

            foreach (HeaderFooter headerFooter in previousSection.HeadersFooters)
                section.HeadersFooters.Add(headerFooter.Clone(true));
        }
        //ExEnd
    }
}
