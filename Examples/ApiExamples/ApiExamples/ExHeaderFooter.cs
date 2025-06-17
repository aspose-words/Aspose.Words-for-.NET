// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
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
        public void Create()
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
            //ExSummary:Shows how to create a header and a footer.
            Document doc = new Document();

            // Create a header and append a paragraph to it. The text in that paragraph
            // will appear at the top of every page of this section, above the main body text.
            HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            doc.FirstSection.HeadersFooters.Add(header);

            Paragraph para = header.AppendParagraph("My header.");

            Assert.That(header.IsHeader, Is.True);
            Assert.That(para.IsEndOfHeaderFooter, Is.True);

            // Create a footer and append a paragraph to it. The text in that paragraph
            // will appear at the bottom of every page of this section, below the main body text.
            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            doc.FirstSection.HeadersFooters.Add(footer);

            para = footer.AppendParagraph("My footer.");

            Assert.That(footer.IsHeader, Is.False);
            Assert.That(para.IsEndOfHeaderFooter, Is.True);

            Assert.That(para.ParentStory, Is.EqualTo(footer));
            Assert.That(para.ParentSection, Is.EqualTo(footer.ParentSection));
            Assert.That(header.ParentSection, Is.EqualTo(footer.ParentSection));

            doc.Save(ArtifactsDir + "HeaderFooter.Create.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "HeaderFooter.Create.docx");

            Assert.That(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].Range.Text
                .Contains("My header."), Is.True);
            Assert.That(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].Range.Text
                .Contains("My footer."), Is.True);
        }

        [Test]
        public void Link()
        {
            //ExStart
            //ExFor:HeaderFooter.IsLinkedToPrevious
            //ExFor:HeaderFooterCollection.Item(Int32)
            //ExFor:HeaderFooterCollection.LinkToPrevious(HeaderFooterType,Boolean)
            //ExFor:HeaderFooterCollection.LinkToPrevious(Boolean)
            //ExFor:HeaderFooter.ParentSection
            //ExSummary:Shows how to link headers and footers between sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 3");

            // Move to the first section and create a header and a footer. By default,
            // the header and the footer will only appear on pages in the section that contains them.
            builder.MoveToSection(0);

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("This is the header, which will be displayed in sections 1 and 2.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("This is the footer, which will be displayed in sections 1, 2 and 3.");

            // We can link a section's headers/footers to the previous section's headers/footers
            // to allow the linking section to display the linked section's headers/footers.
            doc.Sections[1].HeadersFooters.LinkToPrevious(true);

            // Each section will still have its own header/footer objects. When we link sections,
            // the linking section will display the linked section's header/footers while keeping its own.
            Assert.That(doc.Sections[1].HeadersFooters[0], Is.Not.EqualTo(doc.Sections[0].HeadersFooters[0]));
            Assert.That(doc.Sections[1].HeadersFooters[0].ParentSection, Is.Not.EqualTo(doc.Sections[0].HeadersFooters[0].ParentSection));

            // Link the headers/footers of the third section to the headers/footers of the second section.
            // The second section already links to the first section's header/footers,
            // so linking to the second section will create a link chain.
            // The first, second, and now the third sections will all display the first section's headers.
            doc.Sections[2].HeadersFooters.LinkToPrevious(true);

            // We can un-link a previous section's header/footers by passing "false" when calling the LinkToPrevious method.
            doc.Sections[2].HeadersFooters.LinkToPrevious(false);

            // We can also select only a specific type of header/footer to link using this method.
            // The third section now will have the same footer as the second and first sections, but not the header.
            doc.Sections[2].HeadersFooters.LinkToPrevious(HeaderFooterType.FooterPrimary, true);

            // The first section's header/footers cannot link themselves to anything because there is no previous section.
            Assert.That(doc.Sections[0].HeadersFooters.Count, Is.EqualTo(2));
            Assert.That(doc.Sections[0].HeadersFooters.Count(hf => !((HeaderFooter)hf).IsLinkedToPrevious), Is.EqualTo(2));
            
            // All the second section's header/footers are linked to the first section's headers/footers.
            Assert.That(doc.Sections[1].HeadersFooters.Count, Is.EqualTo(6));
            Assert.That(doc.Sections[1].HeadersFooters.Count(hf => ((HeaderFooter)hf).IsLinkedToPrevious), Is.EqualTo(6));

            // In the third section, only the footer is linked to the first section's footer via the second section.
            Assert.That(doc.Sections[2].HeadersFooters.Count, Is.EqualTo(6));
            Assert.That(doc.Sections[2].HeadersFooters.Count(hf => !((HeaderFooter)hf).IsLinkedToPrevious), Is.EqualTo(5));
            Assert.That(doc.Sections[2].HeadersFooters[3].IsLinkedToPrevious, Is.True);

            doc.Save(ArtifactsDir + "HeaderFooter.Link.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "HeaderFooter.Link.docx");

            Assert.That(doc.Sections[0].HeadersFooters.Count, Is.EqualTo(2));
            Assert.That(doc.Sections[0].HeadersFooters.Count(hf => !((HeaderFooter)hf).IsLinkedToPrevious), Is.EqualTo(2));

            Assert.That(doc.Sections[1].HeadersFooters.Count, Is.EqualTo(0));
            Assert.That(doc.Sections[1].HeadersFooters.Count(hf => ((HeaderFooter)hf).IsLinkedToPrevious), Is.EqualTo(0));

            Assert.That(doc.Sections[2].HeadersFooters.Count, Is.EqualTo(5));
            Assert.That(doc.Sections[2].HeadersFooters.Count(hf => !((HeaderFooter)hf).IsLinkedToPrevious), Is.EqualTo(5));
        }

        [Test]
        public void RemoveFooters()
        {
            //ExStart
            //ExFor:Section.HeadersFooters
            //ExFor:HeaderFooterCollection
            //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
            //ExFor:HeaderFooter
            //ExSummary:Shows how to delete all footers from a document.
            Document doc = new Document(MyDir + "Header and footer types.docx");

            // Iterate through each section and remove footers of every kind.
            foreach (Section section in doc.OfType<Section>())
            {
                // There are three kinds of footer and header types.
                // 1 -  The "First" header/footer, which only appears on the first page of a section.
                HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
                footer?.Remove();

                // 2 -  The "Primary" header/footer, which appears on odd pages.
                footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                footer?.Remove();

                // 3 -  The "Even" header/footer, which appears on even pages. 
                footer = section.HeadersFooters[HeaderFooterType.FooterEven];
                footer?.Remove();

                Assert.That(section.HeadersFooters.Count(hf => !((HeaderFooter)hf).IsHeader), Is.EqualTo(0));
            }

            doc.Save(ArtifactsDir + "HeaderFooter.RemoveFooters.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "HeaderFooter.RemoveFooters.docx");

            Assert.That(doc.Sections.Count, Is.EqualTo(1));
            Assert.That(doc.FirstSection.HeadersFooters.Count(hf => !((HeaderFooter)hf).IsHeader), Is.EqualTo(0));
            Assert.That(doc.FirstSection.HeadersFooters.Count(hf => ((HeaderFooter)hf).IsHeader), Is.EqualTo(3));
        }

        [Test]
        public void ExportMode()
        {
            //ExStart
            //ExFor:HtmlSaveOptions.ExportHeadersFootersMode
            //ExFor:ExportHeadersFootersMode
            //ExSummary:Shows how to omit headers/footers when saving a document to HTML.
            Document doc = new Document(MyDir + "Header and footer types.docx");

            // This document contains headers and footers. We can access them via the "HeadersFooters" collection.
            Assert.That(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderFirst].GetText().Trim(), Is.EqualTo("First header"));

            // Formats such as .html do not split the document into pages, so headers/footers will not function the same way
            // they would when we open the document as a .docx using Microsoft Word.
            // If we convert a document with headers/footers to html, the conversion will assimilate the headers/footers into body text.
            // We can use a SaveOptions object to omit headers/footers while converting to html.
            HtmlSaveOptions saveOptions =
                new HtmlSaveOptions(SaveFormat.Html) { ExportHeadersFootersMode = ExportHeadersFootersMode.None };

            doc.Save(ArtifactsDir + "HeaderFooter.ExportMode.html", saveOptions);

            // Open our saved document and verify that it does not contain the header's text
            doc = new Document(ArtifactsDir + "HeaderFooter.ExportMode.html");

            Assert.That(doc.Range.Text.Contains("First header"), Is.False);
            //ExEnd
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
            //ExSummary:Shows how to replace text in a document's footer.
            Document doc = new Document(MyDir + "Footer.docx");

            HeaderFooterCollection headersFooters = doc.FirstSection.HeadersFooters;
            HeaderFooter footer = headersFooters[HeaderFooterType.FooterPrimary];

            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = false
            };

            int currentYear = DateTime.Now.Year;
            footer.Range.Replace("(C) 2006 Aspose Pty Ltd.", $"Copyright (C) {currentYear} by Aspose Pty Ltd.", options);

            doc.Save(ArtifactsDir + "HeaderFooter.ReplaceText.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "HeaderFooter.ReplaceText.docx");

            Assert.That(doc.Range.Text.Contains($"Copyright (C) {currentYear} by Aspose Pty Ltd."), Is.True);
        }

        //ExStart
        //ExFor:IReplacingCallback
        //ExFor:PageSetup.DifferentFirstPageHeaderFooter
        //ExFor:FindReplaceOptions.#ctor(IReplacingCallback)
        //ExSummary:Shows how to track the order in which a text replacement operation traverses nodes.
        [TestCase(false)] //ExSkip
        [TestCase(true)] //ExSkip
        public void Order(bool differentFirstPageHeaderFooter)
        {
            Document doc = new Document(MyDir + "Header and footer types.docx");

            Section firstPageSection = doc.FirstSection;

            ReplaceLog logger = new ReplaceLog();
            FindReplaceOptions options = new FindReplaceOptions(logger);

            // Using a different header/footer for the first page will affect the search order.
            firstPageSection.PageSetup.DifferentFirstPageHeaderFooter = differentFirstPageHeaderFooter;
            doc.Range.Replace(new Regex("(header|footer)"), "", options);

            if (differentFirstPageHeaderFooter)
                Assert.That(logger.Text.Replace("\r", ""), Is.EqualTo("First header\nFirst footer\nSecond header\nSecond footer\nThird header\nThird footer\n"));
            else
                Assert.That(logger.Text.Replace("\r", ""), Is.EqualTo("Third header\nFirst header\nThird footer\nFirst footer\nSecond header\nSecond footer\n"));
        }

        /// <summary>
        /// During a find-and-replace operation, records the contents of every node that has text that the operation 'finds',
        /// in the state it is in before the replacement takes place.
        /// This will display the order in which the text replacement operation traverses nodes.
        /// </summary>
        private class ReplaceLog : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                mTextBuilder.AppendLine(args.MatchNode.GetText());
                return ReplaceAction.Skip;
            }

            internal string Text => mTextBuilder.ToString();

            private readonly StringBuilder mTextBuilder = new StringBuilder();
        }
        //ExEnd
    }
}
