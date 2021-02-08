// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
using Aspose.Words.Fields;
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
            //ExSummary:Shows how to turn off protection for a section.
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Section 1. Hello world!");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.Writeln("Section 2. Hello again!");
            builder.Write("Please enter text here: ");
            builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Placeholder text", 0);

            // Apply write protection to every section in the document.
            doc.Protect(ProtectionType.AllowOnlyFormFields);

            // Turn off write protection for the first section.
            doc.Sections[0].ProtectedForForms = false;

            // In this output document, we will be able to edit the first section freely,
            // and we will only be able to edit the contents of the form field in the second section.
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
            //ExSummary:Shows how to add and remove sections in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");

            Assert.AreEqual("Section 1\x000cSection 2", doc.GetText().Trim());

            // Delete the first section from the document.
            doc.Sections.RemoveAt(0);

            Assert.AreEqual("Section 2", doc.GetText().Trim());
            
            // Append a copy of what is now the first section to the end of the document.
            int lastSectionIdx = doc.Sections.Count - 1;
            Section newSection = doc.Sections[lastSectionIdx].Clone();
            doc.Sections.Add(newSection);

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
            // which contains child nodes that we can edit.
            Assert.AreEqual(1, doc.Sections.Count);

            // Use a document builder to add text to the first section.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Create a second section by inserting a section break.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            Assert.AreEqual(2, doc.Sections.Count);

            // Each section has its own page setup settings.
            // We can split the text in the second section into two columns.
            // This will not affect the text in the first section.
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
        public void CreateManually()
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
            //ExSummary:Shows how to construct an Aspose.Words document by hand.
            Document doc = new Document();

            // A blank document contains one section, one body and one paragraph.
            // Call the "RemoveAllChildren" method to remove all those nodes,
            // and end up with a document node with no children.
            doc.RemoveAllChildren();

            // This document now has no composite child nodes that we can add content to.
            // If we wish to edit it, we will need to repopulate its node collection.
            // First, create a new section, and then append it as a child to the root document node.
            Section section = new Section(doc);
            doc.AppendChild(section);

            // Set some page setup properties for the section.
            section.PageSetup.SectionStart = SectionStart.NewPage;
            section.PageSetup.PaperSize = PaperSize.Letter;
            
            // A section needs a body, which will contain and display all its contents
            // on the page between the section's header and footer.
            Body body = new Body(doc);
            section.AppendChild(body);

            // Create a paragraph, set some formatting properties, and then append it as a child to the body.
            Paragraph para = new Paragraph(doc);

            para.ParagraphFormat.StyleName = "Heading 1";
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            body.AppendChild(para);

            // Finally, add some content to do the document. Create a run,
            // set its appearance and contents, and then append it as a child to the paragraph.
            Run run = new Run(doc);
            run.Text = "Hello World!";
            run.Font.Color = Color.Red;
            para.AppendChild(run);

            Assert.AreEqual("Hello World!", doc.GetText().Trim());

            doc.Save(ArtifactsDir + "Section.CreateManually.docx");
            //ExEnd
        }

        [Test]
        public void EnsureMinimum()
        {
            //ExStart
            //ExFor:NodeCollection.Add
            //ExFor:Section.EnsureMinimum
            //ExFor:SectionCollection.Item(Int32)
            //ExSummary:Shows how to prepare a new section node for editing.
            Document doc = new Document();
            
            // A blank document comes with a section, which has a body, which in turn has a paragraph.
            // We can add contents to this document by adding elements such as text runs, shapes, or tables to that paragraph.
            Assert.AreEqual(NodeType.Section, doc.GetChild(NodeType.Any, 0, true).NodeType);
            Assert.AreEqual(NodeType.Body, doc.Sections[0].GetChild(NodeType.Any, 0, true).NodeType);
            Assert.AreEqual(NodeType.Paragraph, doc.Sections[0].Body.GetChild(NodeType.Any, 0, true).NodeType);

            // If we add a new section like this, it will not have a body, or any other child nodes.
            doc.Sections.Add(new Section(doc));

            Assert.AreEqual(0, doc.Sections[1].GetChildNodes(NodeType.Any, true).Count);

            // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
            doc.LastSection.EnsureMinimum();

            Assert.AreEqual(NodeType.Body, doc.Sections[1].GetChild(NodeType.Any, 0, true).NodeType);
            Assert.AreEqual(NodeType.Paragraph, doc.Sections[1].Body.GetChild(NodeType.Any, 0, true).NodeType);

            doc.Sections[0].Body.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void BodyEnsureMinimum()
        {
            //ExStart
            //ExFor:Section.Body
            //ExFor:Body.EnsureMinimum
            //ExSummary:Clears main text from all sections from the document leaving the sections themselves.
            Document doc = new Document();

            // A blank document contains one section, one body and one paragraph.
            // Call the "RemoveAllChildren" method to remove all those nodes,
            // and end up with a document node with no children.
            doc.RemoveAllChildren();

            // This document now has no composite child nodes that we can add content to.
            // If we wish to edit it, we will need to repopulate its node collection.
            // First, create a new section, and then append it as a child to the root document node.
            Section section = new Section(doc);
            doc.AppendChild(section);

            // A section needs a body, which will contain and display all its contents
            // on the page between the section's header and footer.
            Body body = new Body(doc);
            section.AppendChild(body);

            // This body has no children, so we cannot add runs to it yet.
            Assert.AreEqual(0, doc.FirstSection.Body.GetChildNodes(NodeType.Any, true).Count);
            
            // Call the "EnsureMinimum" to make sure that this body contains at least one empty paragraph. 
            body.EnsureMinimum();

            // Now, we can add runs to the body, and get the document to display them.
            body.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void BodyChildNodes()
        {
            //ExStart
            //ExFor:Body.NodeType
            //ExFor:HeaderFooter.NodeType
            //ExFor:Document.FirstSection
            //ExSummary:Shows how to iterate through the children of a composite node.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Primary header");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Primary footer");

            Section section = doc.FirstSection;

            // A Section is a composite node and can contain child nodes,
            // but only if those child nodes are of a "Body" or "HeaderFooter" node type.
            foreach (Node node in section)
            {
                switch (node.NodeType)
                {
                    case NodeType.Body:
                    {
                        Body body = (Body)node;

                        Console.WriteLine("Body:");
                        Console.WriteLine($"\t\"{body.GetText().Trim()}\"");
                        break;
                    }
                    case NodeType.HeaderFooter:
                    {
                        HeaderFooter headerFooter = (HeaderFooter)node;

                        Console.WriteLine($"HeaderFooter type: {headerFooter.HeaderFooterType}:");
                        Console.WriteLine($"\t\"{headerFooter.GetText().Trim()}\"");
                        break;
                    }
                    default:
                    {
                        throw new Exception("Unexpected node type in a section.");
                    }
                }
            }
            //ExEnd
        }

        [Test]
        public void Clear()
        {
            //ExStart
            //ExFor:NodeCollection.Clear
            //ExSummary:Shows how to remove all sections from a document.
            Document doc = new Document(MyDir + "Document.docx");

            // This document has one section with a few child nodes containing and displaying all the document's contents.
            Assert.AreEqual(1, doc.Sections.Count);
            Assert.AreEqual(19, doc.Sections[0].GetChildNodes(NodeType.Any, true).Count);
            Assert.AreEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.GetText().Trim());

            // Clear the collection of sections, which will remove all of the document's children.
            doc.Sections.Clear();
            
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Any, true).Count);
            Assert.AreEqual(string.Empty, doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void PrependAppendContent()
        {
            //ExStart
            //ExFor:Section.AppendContent
            //ExFor:Section.PrependContent
            //ExSummary:Shows how to append the contents of a section to another section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 3");

            Section section = doc.Sections[2];

            Assert.AreEqual("Section 3" + ControlChar.SectionBreak, section.GetText());

            // Insert the contents of the first section to the beginning of the third section.
            Section sectionToPrepend = doc.Sections[0];
            section.PrependContent(sectionToPrepend);

            // Insert the contents of the second section to the end of the third section.
            Section sectionToAppend = doc.Sections[1];
            section.AppendContent(sectionToAppend);

            // The "PrependContent" and "AppendContent" methods did not create any new sections.
            Assert.AreEqual(3, doc.Sections.Count);
            Assert.AreEqual("Section 1" + ControlChar.ParagraphBreak +
                            "Section 3" + ControlChar.ParagraphBreak +
                            "Section 2" + ControlChar.SectionBreak, section.GetText());
            //ExEnd
        }

        [Test]
        public void ClearContent()
        {
            //ExStart
            //ExFor:Section.ClearContent
            //ExSummary:Shows how to clear the contents of a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
            Assert.AreEqual(1, doc.FirstSection.Body.Paragraphs.Count);

            // Running the "ClearContent" method will remove all the section contents
            // but leave a blank paragraph to add content again.
            doc.FirstSection.ClearContent();

            Assert.AreEqual(string.Empty, doc.GetText().Trim());
            Assert.AreEqual(1, doc.FirstSection.Body.Paragraphs.Count);
            //ExEnd
        }

        [Test]
        public void ClearHeadersFooters()
        {
            //ExStart
            //ExFor:Section.ClearHeadersFooters
            //ExSummary:Shows how to clear the contents of all headers and footers in a section.
            Document doc = new Document();

            Assert.AreEqual(0, doc.FirstSection.HeadersFooters.Count);

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("This is the primary header.");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("This is the primary footer.");

            Assert.AreEqual(2, doc.FirstSection.HeadersFooters.Count);

            Assert.AreEqual("This is the primary header.", doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetText().Trim());
            Assert.AreEqual("This is the primary footer.", doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetText().Trim());

            // Empty all the headers and footers in this section of all their contents.
            // The headers and footers themselves will still be present but will have nothing to display.
            doc.FirstSection.ClearHeadersFooters();

            Assert.AreEqual(2, doc.FirstSection.HeadersFooters.Count);

            Assert.AreEqual(string.Empty, doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetText().Trim());
            Assert.AreEqual(string.Empty, doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetText().Trim());
            //ExEnd
        }

        [Test]
        public void DeleteHeaderFooterShapes()
        {
            //ExStart
            //ExFor:Section.DeleteHeaderFooterShapes
            //ExSummary:Shows how to remove all shapes from all headers footers in a section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a primary header with a shape.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.InsertShape(ShapeType.Rectangle, 100, 100);

            // Create a primary footer with an image.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.InsertImage(ImageDir + "Logo Icon.ico");

            Assert.AreEqual(1, doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetChildNodes(NodeType.Shape, true).Count);
            Assert.AreEqual(1, doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetChildNodes(NodeType.Shape, true).Count);

            // Remove all shapes from the headers and footers in the first section.
            doc.FirstSection.DeleteHeaderFooterShapes();

            Assert.AreEqual(0, doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetChildNodes(NodeType.Shape, true).Count);
            Assert.AreEqual(0, doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetChildNodes(NodeType.Shape, true).Count);
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

            // It is important to understand that a document can contain many sections,
            // and each section has its page setup. In this case, we want to modify them all.
            foreach (Section section in doc.OfType<Section>())
                section.PageSetup.PaperSize = PaperSize.Letter;

            doc.Save(ArtifactsDir + "Section.ModifyPageSetupInAllSections.doc");
        }

        [Test]
        public void CultureInfoPageSetupDefaults()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-us");

            Document docEn = new Document();

            // Assert that page defaults comply with current culture info.
            Section sectionEn = docEn.Sections[0];
            Assert.AreEqual(72.0, sectionEn.PageSetup.LeftMargin); // 2.54 cm         
            Assert.AreEqual(72.0, sectionEn.PageSetup.RightMargin); // 2.54 cm
            Assert.AreEqual(72.0, sectionEn.PageSetup.TopMargin); // 2.54 cm
            Assert.AreEqual(72.0, sectionEn.PageSetup.BottomMargin); // 2.54 cm
            Assert.AreEqual(36.0, sectionEn.PageSetup.HeaderDistance); // 1.27 cm
            Assert.AreEqual(36.0, sectionEn.PageSetup.FooterDistance); // 1.27 cm
            Assert.AreEqual(36.0, sectionEn.PageSetup.TextColumns.Spacing); // 1.27 cm

            // Change the culture and assert that the page defaults are changed.
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

            // Change page defaults.
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