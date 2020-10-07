// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExSection : ApiExampleBase
    {
        [Test]
        public void Protect()
        {
            //ExStart
            //ExFor:Document.Protect(ProtectionType)
            //ExFor:ProtectionType
            //ExFor:Section.ProtectedForForms
            //ExSummary:Shows how to protect a section so only editing in form fields is possible.
            Document doc = new Document();

            // Insert two sections with some text
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Section 1. Unprotected.");
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.Writeln("Section 2. Protected.");

            // Section protection only works when document protection is turned and only editing in form fields is allowed
            doc.Protect(ProtectionType.AllowOnlyFormFields);

            // By default, all sections are protected, but we can selectively turn protection off
            doc.Sections[0].ProtectedForForms = false;

            doc.Save(ArtifactsDir + "Section.Protect.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Section.Protect.docx");

            Assert.False(doc.Sections[0].ProtectedForForms);
            Assert.True(doc.Sections[1].ProtectedForForms);
        }

        [Test]
        public void AddRemove()
        {
            //ExStart
            //ExFor:Document.Sections
            //ExFor:Section.Clone
            //ExFor:SectionCollection
            //ExFor:NodeCollection.RemoveAt(Int32)
            //ExSummary:Shows how to add/remove sections in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");

            // This shows what is in the document originally. The document has two sections
            Assert.AreEqual("Section 1\x000cSection 2", doc.GetText().Trim());

            // Delete the first section from the document
            doc.Sections.RemoveAt(0);

            // Duplicate the last section and append the copy to the end of the document
            int lastSectionIdx = doc.Sections.Count - 1;
            Section newSection = doc.Sections[lastSectionIdx].Clone();
            doc.Sections.Add(newSection);

            // Check what the document contains after we changed it
            Assert.AreEqual("Section 2\x000cSection 2", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void FirstAndLast()
        {
            //ExStart
            //ExFor:Document.FirstSection
            //ExFor:Document.LastSection
            //ExSummary:Shows how to create a new section with a document builder.
            Document doc = new Document();

            // A blank document contains one section by default,
            // in order for us to be able to edit it straight away.
            Assert.AreEqual(1, doc.Sections.Count);

            // Use a document builder to add text, and then to create a new section by inserting a section break.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            Assert.AreEqual(2, doc.Sections.Count);

            // Each section is a subdivision of the document that has its own page setup settings.
            // We can split up the text in the second section into two columns without affecting the first section in any way.
            doc.LastSection.PageSetup.TextColumns.SetCount(2);
            builder.Writeln("Column 1.");
            builder.InsertBreak(BreakType.ColumnBreak);
            builder.Writeln("Column 2.");

            Assert.AreEqual(1, doc.FirstSection.PageSetup.TextColumns.Count);
            Assert.AreEqual(2, doc.LastSection.PageSetup.TextColumns.Count);

            doc.Save(ArtifactsDir + "Section.Create.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Section.Create.docx");

            Assert.AreEqual(1, doc.FirstSection.PageSetup.TextColumns.Count);
            Assert.AreEqual(2, doc.LastSection.PageSetup.TextColumns.Count);
        }

        [Test]
        public void CreateFromScratch()
        {
            //ExStart
            //ExFor:Node.GetText
            //ExFor:CompositeNode.RemoveAllChildren
            //ExFor:CompositeNode.AppendChild
            //ExFor:Section
            //ExFor:Section.#ctor
            //ExFor:Section.PageSetup
            //ExFor:PageSetup.SectionStart
            //ExFor:PageSetup.PaperSize
            //ExFor:SectionStart
            //ExFor:PaperSize
            //ExFor:Body
            //ExFor:Body.#ctor
            //ExFor:Paragraph
            //ExFor:Paragraph.#ctor
            //ExFor:Paragraph.ParagraphFormat
            //ExFor:ParagraphFormat
            //ExFor:ParagraphFormat.StyleName
            //ExFor:ParagraphFormat.Alignment
            //ExFor:ParagraphAlignment
            //ExFor:Run
            //ExFor:Run.#ctor(DocumentBase)
            //ExFor:Run.Text
            //ExFor:Inline.Font
            //ExSummary:Shows how to construct an Aspose Words document node by node.
            Document doc = new Document();

            // A newly created blank document still comes one section, one body and one paragraph
            // Calling this method will remove all those nodes to completely empty the document
            doc.RemoveAllChildren();

            // This document now has no composite nodes that content can be added to
            // If we wish to edit it, we will need to repopulate its node collection,
            // which we will start to do with by creating a new Section node
            Section section = new Section(doc);

            // Append the section to the document
            doc.AppendChild(section);

            // Set some properties for the section
            section.PageSetup.SectionStart = SectionStart.NewPage;
            section.PageSetup.PaperSize = PaperSize.Letter;

            // A section needs a body, which will contain all other nodes that can be edited
            Body body = new Body(doc);
            section.AppendChild(body);

            // The body needs to have at least one paragraph
            // Note that the paragraph has not yet been added to the document, but we have to specify the parent document
            // The parent document is needed so the paragraph can correctly work
            // with styles and other document-wide information
            Paragraph para = new Paragraph(doc);
            body.AppendChild(para);

            // We can set some formatting for the paragraph
            para.ParagraphFormat.StyleName = "Heading 1";
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Now we can begin adding content to the document
            Run run = new Run(doc);
            run.Text = "Hello World!";
            run.Font.Color = Color.Red;
            para.AppendChild(run);

            Assert.AreEqual("Hello World!" + ControlChar.SectionBreakChar, doc.GetText());

            doc.Save(ArtifactsDir + "Section.CreateFromScratch.docx");
            //ExEnd
        }

        [Test]
        public void EnsureSectionMinimum()
        {
            //ExStart
            //ExFor:NodeCollection.Add
            //ExFor:Section.EnsureMinimum
            //ExFor:SectionCollection.Item(Int32)
            //ExSummary:Shows how to prepare a new section node for editing.
            Document doc = new Document();
            
            // A blank document comes with a section, which has a body, which in turn has a paragraph,
            // so we can edit the document by adding children to the paragraph like shapes or runs of text
            Assert.AreEqual(2, doc.Sections[0].GetChildNodes(NodeType.Any, true).Count);

            // If we add a new section like this, it will not have a body or a paragraph that we can edit
            doc.Sections.Add(new Section(doc));

            Assert.AreEqual(0, doc.Sections[1].GetChildNodes(NodeType.Any, true).Count);

            // Makes sure that the section contains a body with at least one paragraph
            doc.LastSection.EnsureMinimum();

            // Now we can add content to this section
            Assert.AreEqual(2, doc.Sections[1].GetChildNodes(NodeType.Any, true).Count);
            //ExEnd
        }

        [Test]
        public void BodyEnsureMinimum()
        {
            //ExStart
            //ExFor:Section.Body
            //ExFor:Body.EnsureMinimum
            //ExSummary:Clears main text from all sections from the document leaving the sections themselves.
            // Open a document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");

            // This shows what is in the document originally
            // The document has two sections
            Assert.AreEqual($"Section 1{ControlChar.SectionBreak}Section 2{ControlChar.SectionBreak}", doc.GetText());

            // Loop through all sections in the document
            foreach (Section section in doc.Sections.OfType<Section>())
            {
                // Each section has a Body node that contains main story (main text) of the section
                Body body = section.Body;

                // This clears all nodes from the body
                body.RemoveAllChildren();

                // Technically speaking, for the main story of a section to be valid, it needs to have
                // at least one empty paragraph. That's what the EnsureMinimum method does
                body.EnsureMinimum();
            }

            // Check how the content of the document looks now
            Assert.AreEqual($"{ControlChar.SectionBreak}{ControlChar.SectionBreak}", doc.GetText());
            //ExEnd
        }

        [Test]
        public void BodyNodeType()
        {
            //ExStart
            //ExFor:Body.NodeType
            //ExFor:HeaderFooter.NodeType
            //ExFor:Document.FirstSection
            //ExSummary:Shows how you can enumerate through children of a composite node and detect types of the children nodes.
            // Open a document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Primary header");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Primary footer");

            // Get the first section in the document
            Section section = doc.FirstSection;

            // A Section is a composite node and therefore can contain child nodes
            // Section can contain only Body and HeaderFooter nodes
            foreach (Node node in section)
            {
                // Every node has the NodeType property
                switch (node.NodeType)
                {
                    case NodeType.Body:
                    {
                        // If the node type is Body, we can cast the node to the Body class
                        Body body = (Body) node;

                        // Write the content of the main story of the section to the console
                        Console.WriteLine("*** Body ***");
                        Console.WriteLine(body.GetText());
                        break;
                    }
                    case NodeType.HeaderFooter:
                    {
                        // If the node type is HeaderFooter, we can cast the node to the HeaderFooter class
                        HeaderFooter headerFooter = (HeaderFooter) node;

                        // Write the content of the header footer to the console
                        Console.WriteLine("*** HeaderFooter ***");
                        Console.WriteLine(headerFooter.HeaderFooterType);
                        Console.WriteLine(headerFooter.GetText());
                        break;
                    }
                    default:
                    {
                        // Other types of nodes never occur inside a Section node
                        throw new Exception("Unexpected node type in a section.");
                    }
                }
            }
            //ExEnd
        }

        [Test]
        public void SectionsDeleteAllSections()
        {
            //ExStart
            //ExFor:NodeCollection.Clear
            //ExSummary:Shows how to remove all sections from a document.
            Document doc = new Document(MyDir + "Document.docx");

            // All the document's content is stored in the child nodes of sections like this one
            Assert.AreEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.GetText().Trim());
            Assert.AreEqual(19, doc.Sections[0].GetChildNodes(NodeType.Any, true).Count);

            doc.Sections.Clear();
            
            // Clearing the section collection effectively empties the document
            Assert.AreEqual(string.Empty, doc.GetText());
            Assert.AreEqual(0, doc.Sections.Count);
            //ExEnd
        }

        [Test]
        public void SectionsAppendSectionContent()
        {
            //ExStart
            //ExFor:Section.AppendContent
            //ExFor:Section.PrependContent
            //ExSummary:Shows how to append content of an existing section. The number of sections in the document remains the same.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 3");

            // This is the section that we will append and prepend to
            Section section = doc.Sections[2];

            // This copies content of the 1st section and inserts it at the beginning of the specified section
            Section sectionToPrepend = doc.Sections[0];
            section.PrependContent(sectionToPrepend);

            // This copies content of the 2nd section and inserts it at the end of the specified section
            Section sectionToAppend = doc.Sections[1];
            section.AppendContent(sectionToAppend);

            Assert.AreEqual("Section 1" + ControlChar.SectionBreak +
                            "Section 2" + ControlChar.SectionBreak +
                            "Section 1" + ControlChar.ParagraphBreak +
                            "Section 3" + ControlChar.ParagraphBreak +
                            "Section 2" + ControlChar.SectionBreak, doc.GetText());
            //ExEnd
        }

        [Test]
        public void SectionsDeleteSectionContent()
        {
            //ExStart
            //ExFor:Section.ClearContent
            //ExSummary:Shows how to clear the content of a section.
            Document doc = new Document(MyDir + "Document.docx");

            Assert.AreEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.GetText().Trim());

            doc.FirstSection.ClearContent();

            Assert.AreEqual(string.Empty, doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void SectionsDeleteHeaderFooter()
        {
            //ExStart
            //ExFor:Section.ClearHeadersFooters
            //ExSummary:Clears content of all headers and footers in a section.
            Document doc = new Document(MyDir + "Header and footer types.docx");

            Section section = doc.Sections[0];

            Assert.AreEqual(6, section.HeadersFooters.Count);
            Assert.AreEqual("First header", section.HeadersFooters[HeaderFooterType.HeaderFirst].GetText().Trim());

            section.ClearHeadersFooters();

            Assert.AreEqual(6, section.HeadersFooters.Count);
            Assert.AreEqual(string.Empty, section.HeadersFooters[HeaderFooterType.HeaderFirst].GetText());
            //ExEnd
        }

        [Test]
        public void SectionDeleteHeaderFooterShapes()
        {
            //ExStart
            //ExFor:Section.DeleteHeaderFooterShapes
            //ExSummary:Removes all images and shapes from all headers footers in a section.
            Document doc = new Document();
            Section section = doc.Sections[0];
            HeaderFooter firstHeader = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);

            section.HeadersFooters.Add(firstHeader);

            firstHeader.AppendParagraph("This paragraph contains a shape: ");

            Shape shape = new Shape(doc, ShapeType.Arrow);
            firstHeader.FirstParagraph.AppendChild(shape);

            Assert.AreEqual(1, firstHeader.GetChildNodes(NodeType.Shape, true).Count);

            section.DeleteHeaderFooterShapes();

            Assert.AreEqual(0, firstHeader.GetChildNodes(NodeType.Shape, true).Count);
            //ExEnd
        }

        [Test]
        public void SectionsCloneSection()
        {
            Document doc = new Document(MyDir + "Document.docx");
            Section cloneSection = doc.Sections[0].Clone();
        }

        [Test]
        public void SectionsImportSection()
        {
            Document srcDoc = new Document(MyDir + "Document.docx");
            Document dstDoc = new Document();

            Section sourceSection = srcDoc.Sections[0];
            Section newSection = (Section) dstDoc.ImportNode(sourceSection, true);
            dstDoc.Sections.Add(newSection);
        }

        [Test]
        public void MigrateFrom2XImportSection()
        {
            Document srcDoc = new Document();
            Document dstDoc = new Document();

            Section sourceSection = srcDoc.Sections[0];
            Section newSection = (Section) dstDoc.ImportNode(sourceSection, true);
            dstDoc.Sections.Add(newSection);
        }

        [Test]
        public void ModifyPageSetupInAllSections()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");

            // It is important to understand that a document can contain many sections and each
            // section has its own page setup. In this case we want to modify them all
            foreach (Section section in doc.OfType<Section>())
                section.PageSetup.PaperSize = PaperSize.Letter;

            doc.Save(ArtifactsDir + "Section.ModifyPageSetupInAllSections.doc");
        }

        [Test]
        public void CultureInfoPageSetupDefaults()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-us");

            Document docEn = new Document();

            // Assert that page defaults comply current culture info
            Section sectionEn = docEn.Sections[0];
            Assert.AreEqual(72.0, sectionEn.PageSetup.LeftMargin); // 2.54 cm         
            Assert.AreEqual(72.0, sectionEn.PageSetup.RightMargin); // 2.54 cm
            Assert.AreEqual(72.0, sectionEn.PageSetup.TopMargin); // 2.54 cm
            Assert.AreEqual(72.0, sectionEn.PageSetup.BottomMargin); // 2.54 cm
            Assert.AreEqual(36.0, sectionEn.PageSetup.HeaderDistance); // 1.27 cm
            Assert.AreEqual(36.0, sectionEn.PageSetup.FooterDistance); // 1.27 cm
            Assert.AreEqual(36.0, sectionEn.PageSetup.TextColumns.Spacing); // 1.27 cm

            // Change culture and assert that the page defaults are changed
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-de");

            Document docDe = new Document();

            Section sectionDe = docDe.Sections[0];
            Assert.AreEqual(70.85, sectionDe.PageSetup.LeftMargin); // 2.5 cm         
            Assert.AreEqual(70.85, sectionDe.PageSetup.RightMargin); // 2.5 cm
            Assert.AreEqual(70.85, sectionDe.PageSetup.TopMargin); // 2.5 cm
            Assert.AreEqual(56.7, sectionDe.PageSetup.BottomMargin); // 2 cm
            Assert.AreEqual(35.4, sectionDe.PageSetup.HeaderDistance); // 1.25 cm
            Assert.AreEqual(35.4, sectionDe.PageSetup.FooterDistance); // 1.25 cm
            Assert.AreEqual(35.4, sectionDe.PageSetup.TextColumns.Spacing); // 1.25 cm

            // Change page defaults
            sectionDe.PageSetup.LeftMargin = 90; // 3.17 cm
            sectionDe.PageSetup.RightMargin = 90; // 3.17 cm
            sectionDe.PageSetup.TopMargin = 72; // 2.54 cm
            sectionDe.PageSetup.BottomMargin = 72; // 2.54 cm
            sectionDe.PageSetup.HeaderDistance = 35.4; // 1.25 cm
            sectionDe.PageSetup.FooterDistance = 35.4; // 1.25 cm
            sectionDe.PageSetup.TextColumns.Spacing = 35.4; // 1.25 cm

            docDe = DocumentHelper.SaveOpen(docDe);

            Section sectionDeAfter = docDe.Sections[0];
            Assert.AreEqual(90.0, sectionDeAfter.PageSetup.LeftMargin); // 3.17 cm         
            Assert.AreEqual(90.0, sectionDeAfter.PageSetup.RightMargin); // 3.17 cm
            Assert.AreEqual(72.0, sectionDeAfter.PageSetup.TopMargin); // 2.54 cm
            Assert.AreEqual(72.0, sectionDeAfter.PageSetup.BottomMargin); // 2.54 cm
            Assert.AreEqual(35.4, sectionDeAfter.PageSetup.HeaderDistance); // 1.25 cm
            Assert.AreEqual(35.4, sectionDeAfter.PageSetup.FooterDistance); // 1.25 cm
            Assert.AreEqual(35.4, sectionDeAfter.PageSetup.TextColumns.Spacing); // 1.25 cm
        }
    }
}