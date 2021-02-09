using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithHeadersAndFooters : DocsExamplesBase
    {
        [Test]
        public void CreateHeaderFooter()
        {
            //ExStart:CreateHeaderFooterUsingDocBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;
            // Specify if we want headers/footers of the first page to be different from other pages.
            // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
            // different headers/footers for odd and even pages.
            pageSetup.DifferentFirstPageHeaderFooter = true;
            pageSetup.HeaderDistance = 20;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;

            builder.Write("Aspose.Words Header/Footer Creation Primer - Title Page.");

            pageSetup.HeaderDistance = 20;
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a positioned image into the top/left corner of the header.
            // Distance from the top/left edges of the page is set to 10 points.
            builder.InsertImage(ImagesDir + "Graphics Interchange Format.gif", RelativeHorizontalPosition.Page, 10,
                RelativeVerticalPosition.Page, 10, 50, 50, WrapType.Through);

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            builder.Write("Aspose.Words Header/Footer Creation Primer.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // We use a table with two cells to make one part of the text on the line (with page numbering).
            // To be aligned left, and the other part of the text (with copyright) to be aligned right.
            builder.StartTable();

            builder.CellFormat.ClearFormatting();

            builder.InsertCell();

            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);

            // It uses PAGE and NUMPAGES fields to auto calculate the current page number and many pages.
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            builder.InsertCell();

            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            builder.Write("(C) 2001 Aspose Pty Ltd. All rights reserved.");

            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            builder.EndRow();
            builder.EndTable();

            builder.MoveToDocumentEnd();

            // Make a page break to create a second page on which the primary headers/footers will be seen.
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            currentSection = builder.CurrentSection;
            pageSetup = currentSection.PageSetup;
            pageSetup.Orientation = Orientation.Landscape;
            // This section does not need a different first-page header/footer we need only one title page in the document,
            // and the header/footer for this page has already been defined in the previous section.
            pageSetup.DifferentFirstPageHeaderFooter = false;

            // This section displays headers/footers from the previous section
            // by default call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this page width
            // is different for the new section, and therefore we need to set different cell widths for a footer table.
            currentSection.HeadersFooters.LinkToPrevious(false);

            // If we want to use the already existing header/footer set for this section.
            // But with some minor modifications, then it may be expedient to copy headers/footers
            // from the previous section and apply the necessary modifications where we want them.
            CopyHeadersFootersFromPreviousSection(currentSection);

            HeaderFooter primaryFooter = currentSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            Row row = primaryFooter.Tables[0].FirstRow;
            row.FirstCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 / 3);
            row.LastCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100 * 2 / 3);

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.CreateHeaderFooter.docx");
            //ExEnd:CreateHeaderFooterUsingDocBuilder
        }

        //ExStart:CopyHeadersFootersFromPreviousSection
        /// <summary>
        /// Clones and copies headers/footers form the previous section to the specified section.
        /// </summary>
        private void CopyHeadersFootersFromPreviousSection(Section section)
        {
            Section previousSection = (Section)section.PreviousSibling;

            if (previousSection == null)
                return;

            section.HeadersFooters.Clear();

            foreach (HeaderFooter headerFooter in previousSection.HeadersFooters)
                section.HeadersFooters.Add(headerFooter.Clone(true));
        }
        //ExEnd:CopyHeadersFootersFromPreviousSection        
    }
}