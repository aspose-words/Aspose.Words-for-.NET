// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class ChangeOrReplaceHeaderAndFooter : TestUtil
    {
        [Test]
        public void CreateHeaderFooter()
        {
            // Create a new Word document.
            using WordprocessingDocument document = WordprocessingDocument.Create(
                ArtifactsDir + "Create header footer - OpenXML.docx",
                WordprocessingDocumentType.Document);

            // Add a main document part.
            MainDocumentPart mainDocumentPart = document.AddMainDocumentPart();
            mainDocumentPart.Document = new Document(new Body());

            // Delete existing header and footer parts (if any).
            mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
            mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

            // Add new header and footer parts.
            HeaderPart headerPart = mainDocumentPart.AddNewPart<HeaderPart>();
            FooterPart footerPart = mainDocumentPart.AddNewPart<FooterPart>();

            string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
            string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);

            // Generate content for the header and footer.
            GenerateHeaderPartContent(headerPart);
            GenerateFooterPartContent(footerPart);

            // Ensure the document has at least one section.
            if (mainDocumentPart.Document.Body.Elements<SectionProperties>().Count() == 0)
            {
                mainDocumentPart.Document.Body.AppendChild(new SectionProperties());
            }

            // Assign the header and footer to all sections.
            IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

            foreach (var section in sections)
            {
                section.RemoveAllChildren<HeaderReference>();
                section.RemoveAllChildren<FooterReference>();

                section.PrependChild(new HeaderReference { Id = headerPartId });
                section.PrependChild(new FooterReference { Id = footerPartId });
            }

            // Save the document.
            mainDocumentPart.Document.Save();
        }

        private void GenerateHeaderPartContent(HeaderPart part)
        {
            Header header = new Header();

            Paragraph paragraph = new Paragraph { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId { Val = "Header" };

            paragraphProperties.Append(paragraphStyleId);

            Run run = new Run();
            Text text = new Text { Text = "Header" };

            run.Append(text);
            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            header.Append(paragraph);

            part.Header = header;
        }

        private void GenerateFooterPartContent(FooterPart part)
        {
            Footer footer = new Footer();

            Paragraph paragraph = new Paragraph { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId { Val = "Footer" };

            paragraphProperties.Append(paragraphStyleId);

            Run run = new Run();
            Text text = new Text { Text = "Footer" };

            run.Append(text);
            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            footer.Append(paragraph);

            part.Footer = footer;
        }
    }
}
