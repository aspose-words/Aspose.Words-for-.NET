using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExBuildingBlocks : ApiExampleBase
    {
        [Test]
        //ExStart
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.#ctor(Aspose.Words.BuildingBlocks.GlossaryDocument) 
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Accept(Aspose.Words.DocumentVisitor)
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Behavior
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Category
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Description
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.FirstSection
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Gallery
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Guid
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.LastSection
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Name
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Sections
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlock.Type
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockBehavior
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockType
        //ExFor:Aspose.Words.BuildingBlocks.NamespaceDoc
        //ExSummary:Shows how to add a custom building block to a document.
        public void BuildingBlocks()
        {
            Document doc = new Document();

            // BuildingBlocks live inside the glossary document.
            // If you're making a document from scratch, the glossary document must also be manually created.
            GlossaryDocument glossaryDoc = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDoc;

            // Create a building block and name it.
            BuildingBlock block = new BuildingBlock(glossaryDoc);
            block.Name = "Custom Block 1";

            Assert.AreEqual("00000000-0000-0000-0000-000000000000", block.Guid.ToString());
            Assert.AreEqual("(Empty Category)", block.Category);
            Assert.AreEqual(BuildingBlockType.None, block.Type);
            Assert.AreEqual(BuildingBlockGallery.All, block.Gallery);
            Assert.AreEqual(BuildingBlockBehavior.Content, block.Behavior);

            glossaryDoc.AppendChild(block);
            Assert.AreEqual(1, glossaryDoc.Count);

            // If we want to use our building block as an AutoText quick part, we need to give it some text and change some properties.
            // All the necessary preparation will be done in a custom document visitor that we will accept. 
            BuildingBlockVisitor visitor = new BuildingBlockVisitor(glossaryDoc);
            block.Accept(visitor);
            
            Console.WriteLine(visitor.GetText());

            doc.Save(MyDir + @"\Artifacts\BuildingBlocks.dotx");
        }

        /// <summary>
        /// Simple implementation of adding text to a building block and preparing it for usage in the document. Implemented as a Visitor.
        /// </summary>
        public class BuildingBlockVisitor : DocumentVisitor
        {
            public BuildingBlockVisitor(GlossaryDocument ownerGlossaryDoc)
            {
                mBuilder = new StringBuilder();
                mGlossaryDoc = ownerGlossaryDoc;
            }

            public string GetText()
            {
                return mBuilder.ToString();
            }

            public override VisitorAction VisitBuildingBlockStart(BuildingBlock block)
            {
                block.Behavior = BuildingBlockBehavior.Paragraph;
                block.Category = "My custom building blocks";
                block.Description = "Using this block in the Quick Parts section of word will place its contents at the cursor.";
                block.Gallery = BuildingBlockGallery.QuickParts;

                // Add content for the block for it to have an effect when used in the document.
                Section section = new Section(mGlossaryDoc);
                block.AppendChild(section);

                Body body = new Body(mGlossaryDoc);
                section.AppendChild(body);

                Paragraph paragraph = new Paragraph(mGlossaryDoc);
                body.AppendChild(paragraph);

                // This run is what will be visible in the document.
                Run run = new Run(mGlossaryDoc, "Text inside " + block.Name);
                block.FirstSection.Body.FirstParagraph.AppendChild(run);

                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBuildingBlockEnd(BuildingBlock block)
            {
                mBuilder.Append("Visited " + block.Name + "\r\n");
                return VisitorAction.Continue;
            }

            private readonly StringBuilder mBuilder;
            private readonly GlossaryDocument mGlossaryDoc;
        }
        //ExEnd

        [Test]
        //ExStart
        //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument
        //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.Accept(Aspose.Words.DocumentVisitor)
        //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.BuildingBlocks
        //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.FirstBuildingBlock
        //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.GetBuildingBlock(Aspose.Words.BuildingBlocks.BuildingBlockGallery,System.String,System.String)
        //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.LastBuildingBlock
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockCollection
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockCollection.Item(System.Int32)
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockCollection.ToArray
        //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockGallery
        //ExSummary:Shows how to use GlossaryDocument and BuildingBlockCollection.
        public void GlossaryDocument()
        {
            Document doc = new Document();
            GlossaryDocument glossaryDoc = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDoc;

            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 1" });
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 2" });
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 3" });
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 4" });
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 5" });

            Assert.AreEqual(5, glossaryDoc.BuildingBlocks.Count);
            Assert.AreEqual("Block 1", glossaryDoc.FirstBuildingBlock.Name);
            Assert.AreEqual("Block 2", glossaryDoc.BuildingBlocks[1].Name);
            Assert.AreEqual("Block 3", glossaryDoc.BuildingBlocks.ToArray()[2].Name);
            Assert.AreEqual("Block 5", glossaryDoc.LastBuildingBlock.Name);

            // Get a block by gallery, category and name.
            BuildingBlock block4 = glossaryDoc.GetBuildingBlock(BuildingBlockGallery.All, "(Empty Category)", "Block 4");

            // All GUIDs are the same by default. To be able to uniquely identify blocks by GUID, these must be changed.
            Assert.AreEqual("00000000-0000-0000-0000-000000000000", block4.Guid.ToString());

            // We will do that using a custom visitor. 
            GlossaryDocVisitor visitor = new GlossaryDocVisitor();
            glossaryDoc.Accept(visitor);

            Assert.AreEqual(5, visitor.GetDictionary().Count);

            Console.WriteLine(visitor.GetText());

            doc.Save(MyDir + @"\Artifacts\GlossaryDocument.dotx");
        }

        /// <summary>
        /// Simple implementation of giving each building block in a glossary document a unique GUID. Implemented as a Visitor.
        /// </summary>
        public class GlossaryDocVisitor : DocumentVisitor
        {
            public GlossaryDocVisitor()
            {
                mBlocks = new Dictionary<Guid, BuildingBlock>();
                mBuilder = new StringBuilder();
            }

            public string GetText()
            {
                return mBuilder.ToString();
            }

            public Dictionary<Guid, BuildingBlock> GetDictionary()
            {
                return mBlocks;
            }

            public override VisitorAction VisitBuildingBlockStart(BuildingBlock block)
            {
                block.Guid = System.Guid.NewGuid();
                mBlocks.Add(block.Guid, block);
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBuildingBlockEnd(BuildingBlock block)
            {
                mBuilder.Append("\tVisited " + block.Name + "\r\n");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitGlossaryDocumentStart(GlossaryDocument glossary)
            {
                mBuilder.Append("Glossary processing started...\r\n");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitGlossaryDocumentEnd(GlossaryDocument glossary)
            {
                mBuilder.Append("Reached end of glossary!\r\nBuildingBlocks found: " + mBlocks.Count);
                return VisitorAction.Continue;
            }

            private readonly Dictionary<System.Guid, BuildingBlock> mBlocks;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd
    }
}
