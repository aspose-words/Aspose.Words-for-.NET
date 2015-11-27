﻿// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Section
{
    [TestFixture]
    public class ExSection : QaTestsBase
    {
        [Test]
        public void Protect()
        {
            //ExStart
            //ExFor:Document.Protect(ProtectionType)
            //ExFor:ProtectionType
            //ExFor:Section.ProtectedForForms
            //ExSummary:Protects a section so only editing in form fields is possible.
            // Create a blank document
            Aspose.Words.Document doc = new Aspose.Words.Document();

            // Insert two sections with some text
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Section 1. Unprotected.");
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.Writeln("Section 2. Protected.");

            // Section protection only works when document protection is turned and only editing in form fields is allowed.
            doc.Protect(ProtectionType.AllowOnlyFormFields);

            // By default, all sections are protected, but we can selectively turn protection off.
            doc.Sections[0].ProtectedForForms = false;

            builder.Document.Save(MyDir + "Section.Protect Out.doc");
            //ExEnd
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
            // Open the document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Section.AddRemove.doc");

            // This shows what is in the document originally. The document has two sections.
            Console.WriteLine(doc.GetText());

            // Delete the first section from the document
            doc.Sections.RemoveAt(0);

            // Duplicate the last section and append the copy to the end of the document.
            int lastSectionIdx = doc.Sections.Count - 1;
            Aspose.Words.Section newSection = doc.Sections[lastSectionIdx].Clone();
            doc.Sections.Add(newSection);

            // Check what the document contains after we changed it.
            Console.WriteLine(doc.GetText());         
            //ExEnd

            Assert.AreEqual("Hello2\x000cHello2\x000c", doc.GetText());
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
            //ExSummary:Creates a simple document from scratch using the Aspose.Words object model.

            // Create an "empty" document. Note that like in Microsoft Word, 
            // the empty document has one section, body and one paragraph in it.
            Aspose.Words.Document doc = new Aspose.Words.Document();

            // This truly makes the document empty. No sections (not possible in Microsoft Word).
            doc.RemoveAllChildren();

            // Create a new section node. 
            // Note that the section has not yet been added to the document, 
            // but we have to specify the parent document.
            Aspose.Words.Section section = new Aspose.Words.Section(doc);

            // Append the section to the document.
            doc.AppendChild(section);

            // Lets set some properties for the section.
            section.PageSetup.SectionStart = SectionStart.NewPage;
            section.PageSetup.PaperSize = PaperSize.Letter;


            // The section that we created is empty, lets populate it. The section needs at least the Body node.
            Body body = new Body(doc);
            section.AppendChild(body);

            
            // The body needs to have at least one paragraph.
            // Note that the paragraph has not yet been added to the document, 
            // but we have to specify the parent document.
            // The parent document is needed so the paragraph can correctly work
            // with styles and other document-wide information.
            Paragraph para = new Paragraph(doc);
            body.AppendChild(para);

            // We can set some formatting for the paragraph
            para.ParagraphFormat.StyleName = "Heading 1";
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;


            // So far we have one empty paragraph in the document.
            // The document is valid and can be saved, but lets add some text before saving.
            // Create a new run of text and add it to our paragraph.
            Run run = new Run(doc);
            run.Text = "Hello World!";
            run.Font.Color = System.Drawing.Color.Red;
            para.AppendChild(run);


            // As a matter of interest, you can retrieve text of the whole document and
            // see that \x000c is automatically appended. \x000c is the end of section character.
            Console.WriteLine("Hello World!\x000c", doc.GetText());

            // Save the document.
            doc.Save(MyDir + "Section.CreateFromScratch Out.doc");
            //ExEnd

            Assert.AreEqual("Hello World!\x000c", doc.GetText());
        }

        [Test]
        public void EnsureSectionMinimum()
        {
            //ExStart
            //ExFor:Section.EnsureMinimum
            //ExSummary:Ensures that a section is valid.
            // Create a blank document
            Aspose.Words.Document doc = new Aspose.Words.Document();
            Aspose.Words.Section section = doc.FirstSection;

            // Makes sure that the section contains a body with at least one paragraph.
            section.EnsureMinimum();
            //ExEnd
        }

        [Test]
        public void BodyEnsureMinimum()
        {
            //ExStart
            //ExFor:Section.Body
            //ExFor:Body.EnsureMinimum
            //ExSummary:Clears main text from all sections from the document leaving the sections themselves.

            // Open a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Section.BodyEnsureMinimum.doc");
            
            // This shows what is in the document originally. The document has two sections.
            Console.WriteLine(doc.GetText());

            // Loop through all sections in the document.
            foreach (Aspose.Words.Section section in doc.Sections)
            {
                // Each section has a Body node that contains main story (main text) of the section.
                Body body = section.Body;

                // This clears all nodes from the body.
                body.RemoveAllChildren();
            
                // Technically speaking, for the main story of a section to be valid, it needs to have
                // at least one empty paragraph. That's what the EnsureMinimum method does.
                body.EnsureMinimum();
            }

            // Check how the content of the document looks now.
            Console.WriteLine(doc.GetText());
            //ExEnd

            Assert.AreEqual("\x000c\x000c", doc.GetText());
        }

        [Test]
        public void BodyNodeType()
        {
            //ExStart
            //ExFor:Body.NodeType
            //ExFor:HeaderFooter.NodeType
            //ExFor:Document.FirstSection
            //ExSummary:Shows how you can enumerate through children of a composite node and detect types of the children nodes.

            // Open a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Section.BodyNodeType.doc");
            
            // Get the first section in the document.
            Aspose.Words.Section section = doc.FirstSection;

            // A Section is a composite node and therefore can contain child nodes.
            // Section can contain only Body and HeaderFooter nodes.
            foreach (Aspose.Words.Node node in section)
            {
                // Every node has the NodeType property.
                switch (node.NodeType)
                {
                    case NodeType.Body:
                    {
                        // If the node type is Body, we can cast the node to the Body class.
                        Body body = (Body)node;

                        // Write the content of the main story of the section to the console.
                        Console.WriteLine("*** Body ***");
                        Console.WriteLine(body.GetText());
                        break;
                    }
                    case NodeType.HeaderFooter:
                    {
                        // If the node type is HeaderFooter, we can cast the node to the HeaderFooter class.
                        Aspose.Words.HeaderFooter headerFooter = (Aspose.Words.HeaderFooter)node;

                        // Write the content of the header footer to the console.
                        Console.WriteLine("*** HeaderFooter ***");
                        Console.WriteLine(headerFooter.HeaderFooterType);
                        Console.WriteLine(headerFooter.GetText());
                        break;
                    }
                    default:
                    {
                        // Other types of nodes never occur inside a Section node.
                        throw new Exception("Unexpected node type in a section.");
                    }
                }
            }
            //ExEnd
        }

        [Test]
        public void SectionsAccessByIndex()
        {
            //ExStart
            //ExFor:SectionCollection.Item(Int32)
            //ExId:SectionsAccessByIndex
            //ExSummary:Shows how to access a section at the specified index.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Section section = doc.Sections[0];
            //ExEnd
        }

        [Test]
        public void SectionsAddSection()
        {
            //ExStart
            //ExFor:NodeCollection.Add
            //ExId:SectionsAddSection
            //ExSummary:Shows how to add a section to the end of the document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Section sectionToAdd = new Aspose.Words.Section(doc); 
            doc.Sections.Add(sectionToAdd);
            //ExEnd
        }

        [Test]
        public void SectionsDeleteSection()
        {
            //ExStart
            //ExId:SectionsDeleteSection
            //ExSummary:Shows how to remove a section at the specified index.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            doc.Sections.RemoveAt(0);
            //ExEnd
        }

        [Test]
        public void SectionsDeleteAllSections()
        {
            //ExStart
            //ExFor:NodeCollection.Clear
            //ExId:SectionsDeleteAllSections
            //ExSummary:Shows how to remove all sections from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            doc.Sections.Clear();
            //ExEnd
        }

        [Test]
        public void SectionsAppendSectionContent()
        {
            //ExStart
            //ExFor:Section.AppendContent
            //ExFor:Section.PrependContent
            //ExId:SectionsAppendSectionContent
            //ExSummary:Shows how to append content of an existing section. The number of sections in the document remains the same.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Section.AppendContent.doc");
            
            // This is the section that we will append and prepend to.
            Aspose.Words.Section section = doc.Sections[2];

            // This copies content of the 1st section and inserts it at the beginning of the specified section.
            Aspose.Words.Section sectionToPrepend = doc.Sections[0];
            section.PrependContent(sectionToPrepend);

            // This copies content of the 2nd section and inserts it at the end of the specified section.
            Aspose.Words.Section sectionToAppend = doc.Sections[1];
            section.AppendContent(sectionToAppend);
            //ExEnd
        }

        [Test]
        public void SectionsDeleteSectionContent()
        {
            //ExStart
            //ExFor:Section.ClearContent
            //ExId:SectionsDeleteSectionContent
            //ExSummary:Shows how to delete main content of a section.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Section section = doc.Sections[0];
            section.ClearContent();
            //ExEnd
        }

        [Test]
        public void SectionsDeleteHeaderFooter()
        {
            //ExStart
            //ExFor:Section.ClearHeadersFooters
            //ExId:SectionsDeleteHeaderFooter
            //ExSummary:Clears content of all headers and footers in a section.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Section section = doc.Sections[0];
            section.ClearHeadersFooters();
            //ExEnd
        }

        [Test]
        public void SectionDeleteHeaderFooterShapes()
        {
            //ExStart
            //ExFor:Section.DeleteHeaderFooterShapes
            //ExSummary:Removes all images and shapes from all headers footers in a section.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Section section = doc.Sections[0];
            section.DeleteHeaderFooterShapes();
            //ExEnd
        }


        [Test]
        public void SectionsCloneSection()
        {
            //ExStart
            //ExId:SectionsCloneSection
            //ExSummary:Shows how to create a duplicate of a particular section.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Section cloneSection = doc.Sections[0].Clone();
            //ExEnd
        }

        [Test]
        public void SectionsImportSection()
        {
            //ExStart
            //ExId:SectionsImportSection
            //ExSummary:Shows how to copy sections between documents.
            Aspose.Words.Document srcDoc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Document dstDoc = new Aspose.Words.Document();

            Aspose.Words.Section sourceSection = srcDoc.Sections[0];
            Aspose.Words.Section newSection = (Aspose.Words.Section)dstDoc.ImportNode(sourceSection, true);
            dstDoc.Sections.Add(newSection);
            //ExEnd
        }

        [Test]
        public void MigrateFrom2XImportSection()
        {
            Aspose.Words.Document srcDoc = new Aspose.Words.Document();
            Aspose.Words.Document dstDoc = new Aspose.Words.Document();

            //ExStart
            //ExId:MigrateFrom2XImportSection
            //ExSummary:This fragment shows how to insert a section from another document in Aspose.Words 3.0 or higher.
            Aspose.Words.Section sourceSection = srcDoc.Sections[0];
            Aspose.Words.Section newSection = (Aspose.Words.Section)dstDoc.ImportNode(sourceSection, true);
            dstDoc.Sections.Add(newSection);
            //ExEnd
        }

        [Test]
        public void ModifyPageSetupInAllSections()
        {
            //ExStart
            //ExId:ModifyPageSetupInAllSections
            //ExSummary:Shows how to set paper size for the whole document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Section.ModifyPageSetupInAllSections.doc");

            // It is important to understand that a document can contain many sections and each
            // section has its own page setup. In this case we want to modify them all.
            foreach (Aspose.Words.Section section in doc)
                section.PageSetup.PaperSize = PaperSize.Letter;

            doc.Save(MyDir + "Section.ModifyPageSetupInAllSections Out.doc");
            //ExEnd
        }
    }
}
