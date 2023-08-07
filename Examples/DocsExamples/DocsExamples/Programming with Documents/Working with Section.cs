using Aspose.Words;
using NUnit.Framework;
using System;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithSection : DocsExamplesBase
    {
        [Test]
        public void AddSection()
        {
            //ExStart:AddSection
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello1");
            builder.Writeln("Hello2");

            Section sectionToAdd = new Section(doc);
            doc.Sections.Add(sectionToAdd);
            //ExEnd:AddSection
        }

        [Test]
        public void DeleteSection()
        {
            //ExStart:DeleteSection
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello1");
            doc.AppendChild(new Section(doc));
            builder.Writeln("Hello2");
            doc.AppendChild(new Section(doc));

            doc.Sections.RemoveAt(0);
            //ExEnd:DeleteSection
        }

        [Test]
        public void DeleteAllSections()
        {
            //ExStart:DeleteAllSections
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello1");
            doc.AppendChild(new Section(doc));
            builder.Writeln("Hello2");
            doc.AppendChild(new Section(doc));

            doc.Sections.Clear();
            //ExEnd:DeleteAllSections
        }

        [Test]
        public void AppendSectionContent()
        {
            //ExStart:AppendSectionContent
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 3");

            Section section = doc.Sections[2];

            // Insert the contents of the first section to the beginning of the third section.
            Section sectionToPrepend = doc.Sections[0];
            section.PrependContent(sectionToPrepend);

            // Insert the contents of the second section to the end of the third section.
            Section sectionToAppend = doc.Sections[1];
            section.AppendContent(sectionToAppend);
            //ExEnd:AppendSectionContent
        }

        [Test]
        public void CloneSection()
        {
            //ExStart:CloneSection
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
            Document doc = new Document(MyDir + "Document.docx");
            Section cloneSection = doc.Sections[0].Clone();
            //ExEnd:CloneSection
        }

        [Test]
        public void CopySection()
        {
            //ExStart:CopySection
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
            Document srcDoc = new Document(MyDir + "Document.docx");
            Document dstDoc = new Document();

            Section sourceSection = srcDoc.Sections[0];
            Section newSection = (Section)dstDoc.ImportNode(sourceSection, true);
            dstDoc.Sections.Add(newSection);

            dstDoc.Save(ArtifactsDir + "WorkingWithSection.CopySection.docx");
            //ExEnd:CopySection
        }

        [Test]
        public void DeleteHeaderFooterContent()
        {
            //ExStart:DeleteHeaderFooterContent
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
            Document doc = new Document(MyDir + "Document.docx");

            Section section = doc.Sections[0];
            section.ClearHeadersFooters();
            //ExEnd:DeleteHeaderFooterContent
        }

        [Test]
        public void DeleteHeaderFooterShapes()
        {
            //ExStart:DeleteHeaderFooterShapes
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
            Document doc = new Document(MyDir + "Document.docx");

            Section section = doc.Sections[0];
            section.DeleteHeaderFooterShapes();
            //ExEnd:DeleteHeaderFooterShapes
        }

        [Test]
        public void DeleteSectionContent()
        {
            //ExStart:DeleteSectionContent
            Document doc = new Document(MyDir + "Document.docx");

            Section section = doc.Sections[0];
            section.ClearContent();
            //ExEnd:DeleteSectionContent
        }

        [Test]
        public void ModifyPageSetupInAllSections()
        {
            //ExStart:ModifyPageSetupInAllSections
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Section 1");
            doc.AppendChild(new Section(doc));
            builder.Writeln("Section 2");
            doc.AppendChild(new Section(doc));
            builder.Writeln("Section 3");
            doc.AppendChild(new Section(doc));
            builder.Writeln("Section 4");

            // It is important to understand that a document can contain many sections,
            // and each section has its page setup. In this case, we want to modify them all.
            foreach (Section section in doc)
                section.PageSetup.PaperSize = PaperSize.Letter;

            doc.Save(ArtifactsDir + "WorkingWithSection.ModifyPageSetupInAllSections.doc");
            //ExEnd:ModifyPageSetupInAllSections
        }

        [Test]
        public void SectionsAccessByIndex()
        {
            //ExStart:SectionsAccessByIndex
            Document doc = new Document(MyDir + "Document.docx");

            Section section = doc.Sections[0];
            section.PageSetup.LeftMargin = 90; // 3.17 cm
            section.PageSetup.RightMargin = 90; // 3.17 cm
            section.PageSetup.TopMargin = 72; // 2.54 cm
            section.PageSetup.BottomMargin = 72; // 2.54 cm
            section.PageSetup.HeaderDistance = 35.4; // 1.25 cm
            section.PageSetup.FooterDistance = 35.4; // 1.25 cm
            section.PageSetup.TextColumns.Spacing = 35.4; // 1.25 cm
            //ExEnd:SectionsAccessByIndex
        }

        [Test]
        public void SectionChildNodes()
        {
            //ExStart:SectionChildNodes
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
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
            //ExEnd:SectionChildNodes
        }

        [Test]
        public void EnsureMinimum()
        {
            //ExStart:EnsureMinimum
            //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
            Document doc = new Document();

            // If we add a new section like this, it will not have a body, or any other child nodes.
            doc.Sections.Add(new Section(doc));
            // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
            doc.LastSection.EnsureMinimum();
            
            doc.Sections[0].Body.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));
            //ExEnd:EnsureMinimum
        }
    }
}