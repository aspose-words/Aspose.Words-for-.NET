// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExHeaderFooter : ApiExampleBase
    {
        [Test]
        public void HeaderFooterCreate()
        {
            //ExStart
            //ExFor:HeaderFooter
            //ExFor:HeaderFooter.#ctor(DocumentBase, HeaderFooterType)
            //ExFor:HeaderFooter.HeaderFooterType
            //ExFor:HeaderFooter.IsHeader
            //ExFor:HeaderFooterCollection
            //ExFor:Paragraph.IsEndOfHeaderFooter
            //ExFor:Paragraph.ParentSection
            //ExFor:Paragraph.ParentStory
            //ExFor:Story.AppendParagraph
            //ExSummary:Creates a header and footer using the document object model and insert them into a section.
            Document doc = new Document();
            
            HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            doc.FirstSection.HeadersFooters.Add(header);

            // Add a paragraph with text to the footer.
            Paragraph para = header.AppendParagraph("My header");

            Assert.True(header.IsHeader);
            Assert.True(para.IsEndOfHeaderFooter);

            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            doc.FirstSection.HeadersFooters.Add(footer);

            // Add a paragraph with text to the footer.
            para = footer.AppendParagraph("My footer");

            Assert.False(footer.IsHeader);
            Assert.True(para.IsEndOfHeaderFooter);

            Assert.AreEqual(footer, para.ParentStory);
            Assert.AreEqual(footer.ParentSection, para.ParentSection);
            Assert.AreEqual(footer.ParentSection, header.ParentSection);
            
            doc.Save(ArtifactsDir + "HeaderFooter.HeaderFooterCreate.docx");
            //ExEnd
            doc = new Document(ArtifactsDir + "HeaderFooter.HeaderFooterCreate.docx");

            Assert.True(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].Range.Text
                .Contains("My header"));
            Assert.True(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].Range.Text
                .Contains("My footer"));
        }

        [Test]
        public void HeaderFooterLink()
        {
            //ExStart
            //ExFor:HeaderFooter.IsLinkedToPrevious
            //ExFor:HeaderFooterCollection.Item(System.Int32)
            //ExFor:HeaderFooterCollection.LinkToPrevious(Aspose.Words.HeaderFooterType,System.Boolean)
            //ExFor:HeaderFooterCollection.LinkToPrevious(System.Boolean)
            //ExFor:HeaderFooter.ParentSection
            //ExSummary:Shows how to link header/footers between sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create three sections
            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 3");

            // Create a header and footer in the first section and give them text
            builder.MoveToSection(0);

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("This is the header, which will be displayed in sections 1 and 2.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("This is the footer, which will be displayed in sections 1, 2 and 3.");

            // If headers/footers are linked by the next section, they appear in that section also
            // The second section will display the header/footers of the first
            doc.Sections[1].HeadersFooters.LinkToPrevious(true);

            // However, the underlying headers/footers in the respective header/footer collections of the sections still remain different
            // Linking just overrides the existing headers/footers from the latter section
            Assert.AreEqual(doc.Sections[0].HeadersFooters[0].HeaderFooterType, doc.Sections[1].HeadersFooters[0].HeaderFooterType);
            Assert.AreNotEqual(doc.Sections[0].HeadersFooters[0].ParentSection, doc.Sections[1].HeadersFooters[0].ParentSection);
            Assert.AreNotEqual(doc.Sections[0].HeadersFooters[0].GetText(), doc.Sections[1].HeadersFooters[0].GetText());

            // Likewise, unlinking headers/footers makes them not appear
            doc.Sections[2].HeadersFooters.LinkToPrevious(false);

            // We can also choose only certain header/footer types to get linked, like the footer in this case
            // The 3rd section now won't have the same header but will have the same footer as the 2nd and 1st sections
            doc.Sections[2].HeadersFooters.LinkToPrevious(HeaderFooterType.FooterPrimary, true);
            
            // The first section's header/footers can't link themselves to anything because there is no previous section
            Assert.AreEqual(2, doc.Sections[0].HeadersFooters.Count);
            Assert.False(doc.Sections[0].HeadersFooters[0].IsLinkedToPrevious);
            Assert.False(doc.Sections[0].HeadersFooters[1].IsLinkedToPrevious);

            // All of the second section's header/footers are linked to those of the first
            Assert.AreEqual(6, doc.Sections[1].HeadersFooters.Count);
            Assert.True(doc.Sections[1].HeadersFooters[0].IsLinkedToPrevious);
            Assert.True(doc.Sections[1].HeadersFooters[1].IsLinkedToPrevious);
            Assert.True(doc.Sections[1].HeadersFooters[2].IsLinkedToPrevious);
            Assert.True(doc.Sections[1].HeadersFooters[3].IsLinkedToPrevious);
            Assert.True(doc.Sections[1].HeadersFooters[4].IsLinkedToPrevious);
            Assert.True(doc.Sections[1].HeadersFooters[5].IsLinkedToPrevious);

            // In the third section, only the footer we explicitly linked is linked to that of the second, and consequently the first section
            Assert.AreEqual(6, doc.Sections[2].HeadersFooters.Count);
            Assert.False(doc.Sections[2].HeadersFooters[0].IsLinkedToPrevious);
            Assert.False(doc.Sections[2].HeadersFooters[1].IsLinkedToPrevious);
            Assert.False(doc.Sections[2].HeadersFooters[2].IsLinkedToPrevious);
            Assert.True(doc.Sections[2].HeadersFooters[3].IsLinkedToPrevious);
            Assert.False(doc.Sections[2].HeadersFooters[4].IsLinkedToPrevious);
            Assert.False(doc.Sections[2].HeadersFooters[5].IsLinkedToPrevious);
    
            doc.Save(ArtifactsDir + "HeaderFooter.HeaderFooterLink.docx");
            //ExEnd
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
            //ExSummary:Deletes all footers from all sections, but leaves headers intact.
            Document doc = new Document(MyDir + "HeaderFooter.RemoveFooters.doc");

            foreach (Section section in doc.OfType<Section>())
            {
                // Up to three different footers are possible in a section (for first, even and odd pages).
                // We check and delete all of them.
                HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
                footer?.Remove();

                // Primary footer is the footer used for odd pages.
                footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                footer?.Remove();

                footer = section.HeadersFooters[HeaderFooterType.FooterEven];
                footer?.Remove();
            }

            doc.Save(ArtifactsDir + "HeaderFooter.RemoveFooters.doc");
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

            // Disables exporting headers and footers.
            HtmlSaveOptions saveOptions =
                new HtmlSaveOptions(SaveFormat.Html) { ExportHeadersFootersMode = ExportHeadersFootersMode.None };

            doc.Save(ArtifactsDir + "HeaderFooter.DisableHeadersFooters.html", saveOptions);
            //ExEnd

            // Verify that the output document is correct.
            doc = new Document(ArtifactsDir + "HeaderFooter.DisableHeadersFooters.html");
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
            //ExFor:Range.Replace(String, String, FindReplaceOptions)
            //ExSummary:Shows how to replace text in the document footer.
            // Open the template document, containing obsolete copyright information in the footer.
            Document doc = new Document(MyDir + "HeaderFooter.ReplaceText.doc");

            HeaderFooterCollection headersFooters = doc.FirstSection.HeadersFooters;
            HeaderFooter footer = headersFooters[HeaderFooterType.FooterPrimary];

            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = false
            };

            footer.Range.Replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2011 by Aspose Pty Ltd.", options);

            doc.Save(ArtifactsDir + "HeaderFooter.ReplaceText.doc");
            //ExEnd

            // Verify that the appropriate changes were made to the output document.
            doc = new Document(ArtifactsDir + "HeaderFooter.ReplaceText.doc");
            Assert.IsTrue(doc.Range.Text.Contains("Copyright (C) 2011 by Aspose Pty Ltd."));
        }

        //ExStart
        //ExFor:IReplacingCallback
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExSummary:Show changes for headers and footers order.
        [Test] //ExSkip
        public void HeaderFooterOrder()
        {            
            Document doc = new Document(MyDir + "HeaderFooter.HeaderFooterOrder.docx");

            // Assert that we use special header and footer for the first page
            // The order for this: first header\footer, even header\footer, primary header\footer
            Section firstPageSection = doc.FirstSection;
            Assert.AreEqual(true, firstPageSection.PageSetup.DifferentFirstPageHeaderFooter);

            ReplaceLog logger = new ReplaceLog();
            FindReplaceOptions options = new FindReplaceOptions { ReplacingCallback = logger };
            doc.Range.Replace(new Regex("(header|footer)"), "", options);

            doc.Save(ArtifactsDir + "HeaderFooter.HeaderFooterOrder.docx");
            
            #if NETFRAMEWORK || NETSTANDARD2_0
            Assert.AreEqual("First header\nFirst footer\nSecond header\nSecond footer\nThird header\n" +
                "Third footer\n", logger.Text.Replace("\r", ""));            
            #else
            Assert.AreEqual("First header\nFirst footer\nSecond header\nSecond footer\nThird header\n" +
                "Third footer\n", logger.Text);
            #endif
            
            // Prepare our string builder for assert results without "DifferentFirstPageHeaderFooter"
            logger.ClearText();

            // Remove special first page
            // The order for this: primary header, default header, primary footer, default footer, even header\footer
            firstPageSection.PageSetup.DifferentFirstPageHeaderFooter = false;
            doc.Range.Replace(new Regex("(header|footer)"), "", options);
            
            #if NETFRAMEWORK || NETSTANDARD2_0
            Assert.AreEqual("Third header\nFirst header\nThird footer\nFirst footer\nSecond header\n" +
                "Second footer\n", logger.Text.Replace("\r", ""));
            #else
            Assert.AreEqual("Third header\nFirst header\nThird footer\nFirst footer\nSecond header\n" +
                "Second footer\n", logger.Text);
            #endif
        }

        private class ReplaceLog : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                _textBuilder.AppendLine(args.MatchNode.GetText());
                return ReplaceAction.Skip;
            }

            internal void ClearText()
            {
                _textBuilder.Clear();
            }

            internal string Text
            {
                get { return _textBuilder.ToString(); }
            }

            private readonly StringBuilder _textBuilder = new StringBuilder();
        }
        //ExEnd

        [Test]
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
            string imageFileName = ImageDir + "Aspose.Words.gif";
            builder.InsertImage(imageFileName, RelativeHorizontalPosition.Page, 10, RelativeVerticalPosition.Page, 10,
                50, 50, WrapType.Through);

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
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100.0F / 3);

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
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100.0F * 2 / 3);

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
            row.FirstCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100.0F / 3);
            row.LastCell.CellFormat.PreferredWidth = PreferredWidth.FromPercent(100.0F * 2 / 3);

            // Save the resulting document.
            doc.Save(ArtifactsDir + "HeaderFooter.Primer.doc");
        }

        /// <summary>
        /// Clones and copies headers/footers form the previous section to the specified section.
        /// </summary>
        private static void CopyHeadersFootersFromPreviousSection(Section section)
        {
            Section previousSection = (Section) section.PreviousSibling;

            if (previousSection == null)
                return;

            section.HeadersFooters.Clear();

            foreach (HeaderFooter headerFooter in previousSection.HeadersFooters.OfType<HeaderFooter>())
            {
                section.HeadersFooters.Add(headerFooter.Clone(true));
            }
        }
    }
}