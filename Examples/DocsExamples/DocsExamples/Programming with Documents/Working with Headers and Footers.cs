using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithHeadersAndFooters : DocsExamplesBase
    {
        [Test]
        public void CreateHeaderFooter()
        {
            //ExStart:CreateHeaderFooter
            //GistId:84cab3a22008f041ee6c1e959da09949
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //ExStart:HeaderFooterType
            //GistId:84cab3a22008f041ee6c1e959da09949
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header for the first page.");
            //ExEnd:HeaderFooterType

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Header for odd page.");

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.CreateHeaderFooter.docx");
            //ExEnd:CreateHeaderFooter
        }

        [Test]
        public void DifferentFirstPage()
        {
            //ExStart:DifferentFirstPage
            //GistId:84cab3a22008f041ee6c1e959da09949
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify that we want different headers and footers for first page.
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header for the first page.");

            builder.MoveToSection(0);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.DifferentFirstPage.docx");
            //ExEnd:DifferentFirstPage
        }

        [Test]
        public void OddEvenPages()
        {
            //ExStart:OddEvenPages
            //GistId:84cab3a22008f041ee6c1e959da09949
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Specify that we want different headers and footers for even and odd pages.            
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write("Header for even pages.");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header for odd pages.");

            builder.MoveToSection(0);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.OddEvenPages.docx");
            //ExEnd:OddEvenPages
        }

        [Test]
        public void InsertImage()
        {
            //ExStart:InsertImage
            //GistId:84cab3a22008f041ee6c1e959da09949
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);            
            builder.InsertImage(ImagesDir + "Logo.jpg", RelativeHorizontalPosition.RightMargin, 10,
                RelativeVerticalPosition.Page, 10, 50, 50, WrapType.Through);            

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.InsertImage.docx");
            //ExEnd:InsertImage
        }

        [Test]
        public void FontProps()
        {
            //ExStart:FontProps
            //GistId:84cab3a22008f041ee6c1e959da09949
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            builder.Write("Header for odd page.");

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.HeaderFooterFontProps.docx");
            //ExEnd:FontProps
        }

        [Test]
        public void PageNumbers()
        {
            //ExStart:PageNumbers
            //GistId:84cab3a22008f041ee6c1e959da09949
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.PageNumbers.docx");
            //ExEnd:PageNumbers
        }

        [Test]
        public void LinkToPreviousHeaderFooter()
        {
            //ExStart:LinkToPreviousHeaderFooter
            //GistId:84cab3a22008f041ee6c1e959da09949
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.PageSetup.DifferentFirstPageHeaderFooter = true;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            builder.Write("Header for the first page.");

            builder.MoveToDocumentEnd();            
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            Section currentSection = builder.CurrentSection;
            PageSetup pageSetup = currentSection.PageSetup;
            pageSetup.Orientation = Orientation.Landscape;
            // This section does not need a different first-page header/footer we need only one title page in the document,
            // and the header/footer for this page has already been defined in the previous section.
            pageSetup.DifferentFirstPageHeaderFooter = false;

            // This section displays headers/footers from the previous section
            // by default call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this page width
            // is different for the new section.
            currentSection.HeadersFooters.LinkToPrevious(false);
            currentSection.HeadersFooters.Clear();

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            builder.Font.Name = "Arial";            
            builder.Font.Size = 12;
            builder.Write("New Header for the first page.");

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.LinkToPreviousHeaderFooter.docx");
            //ExEnd:LinkToPreviousHeaderFooter
        }

        [Test]
        public void SectionsWithDifferentHeaders()
        {
            //ExStart:SectionsWithDifferentHeaders            
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d            
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            PageSetup pageSetup = builder.CurrentSection.PageSetup;
            pageSetup.DifferentFirstPageHeaderFooter = true;
            pageSetup.HeaderDistance = 20;

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Font.Size = 14;
            builder.Write("Header for the first page.");
                        
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            // Insert a positioned image into the top/left corner of the header.
            // Distance from the top/left edges of the page is set to 10 points.
            builder.InsertImage(ImagesDir + "Logo.jpg", RelativeHorizontalPosition.Page, 10,
                RelativeVerticalPosition.Page, 10, 50, 50, WrapType.Through);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Header for odd page.");            

            doc.Save(ArtifactsDir + "WorkingWithHeadersAndFooters.SectionsWithDifferentHeaders.docx");
            //ExEnd:SectionsWithDifferentHeaders
        }

        //ExStart:CopyHeadersFootersFromPreviousSection
        //GistId:84cab3a22008f041ee6c1e959da09949
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