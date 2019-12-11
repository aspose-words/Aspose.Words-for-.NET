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
        //ExStart
        //ExFor:Document.GlossaryDocument
        //ExFor:BuildingBlocks.BuildingBlock
        //ExFor:BuildingBlocks.BuildingBlock.#ctor(BuildingBlocks.GlossaryDocument) 
        //ExFor:BuildingBlocks.BuildingBlock.Accept(DocumentVisitor)
        //ExFor:BuildingBlocks.BuildingBlock.Behavior
        //ExFor:BuildingBlocks.BuildingBlock.Category
        //ExFor:BuildingBlocks.BuildingBlock.Description
        //ExFor:BuildingBlocks.BuildingBlock.FirstSection
        //ExFor:BuildingBlocks.BuildingBlock.Gallery
        //ExFor:BuildingBlocks.BuildingBlock.Guid
        //ExFor:BuildingBlocks.BuildingBlock.LastSection
        //ExFor:BuildingBlocks.BuildingBlock.Name
        //ExFor:BuildingBlocks.BuildingBlock.Sections
        //ExFor:BuildingBlocks.BuildingBlock.Type
        //ExFor:BuildingBlocks.BuildingBlockBehavior
        //ExFor:BuildingBlocks.BuildingBlockType
        //ExSummary:Shows how to add a custom building block to a document.
        [Test] //ExSkip
        public void BuildingBlockFields()
        {
            Document doc = new Document();

            // BuildingBlocks live inside the glossary document
            // If you're making a document from scratch, the glossary document must also be manually created
            GlossaryDocument glossaryDoc = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDoc;

            // Create a building block and name it
            BuildingBlock block = new BuildingBlock(glossaryDoc);
            block.Name = "Custom Block";
            
            // Put in in the document's glossary document
            glossaryDoc.AppendChild(block);
            Assert.AreEqual(1, glossaryDoc.Count);

            // All GUIDs are this value by default
            Assert.AreEqual("00000000-0000-0000-0000-000000000000", block.Guid.ToString());

            // In Microsoft Word, we can use these attributes to find blocks in Insert > Quick Parts > Building Blocks Organizer  
            Assert.AreEqual("(Empty Category)", block.Category);
            Assert.AreEqual(BuildingBlockType.None, block.Type);
            Assert.AreEqual(BuildingBlockGallery.All, block.Gallery);
            Assert.AreEqual(BuildingBlockBehavior.Content, block.Behavior);

            // If we want to use our building block as an AutoText quick part, we need to give it some text and change some properties
            // All the necessary preparation will be done in a custom document visitor that we will accept
            BuildingBlockVisitor visitor = new BuildingBlockVisitor(glossaryDoc);
            block.Accept(visitor);

            // We can find the block we made in the glossary document like this
            BuildingBlock customBlock = glossaryDoc.GetBuildingBlock(BuildingBlockGallery.QuickParts,
                "My custom building blocks", "Custom Block");

            // Our block contains one section which now contains our text
            Assert.AreEqual("Text inside " + customBlock.Name + '\f',
                customBlock.FirstSection.Body.FirstParagraph.GetText());
            Assert.AreEqual(customBlock.FirstSection, customBlock.LastSection);

            Assert.AreNotEqual("00000000-0000-0000-0000-000000000000", customBlock.Guid.ToString());
            Assert.AreEqual("My custom building blocks", customBlock.Category);
            Assert.AreEqual(BuildingBlockType.None, customBlock.Type);
            Assert.AreEqual(BuildingBlockGallery.QuickParts, customBlock.Gallery);
            Assert.AreEqual(BuildingBlockBehavior.Paragraph, customBlock.Behavior);

            // Then we can insert it into the document as a new section
            doc.AppendChild(doc.ImportNode(customBlock.FirstSection, true));

            // Or we can find it in Microsoft Word's Building Blocks Organizer and place it manually
            doc.Save(ArtifactsDir + "BuildingBlocks.BuildingBlockFields.dotx");
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

            public override VisitorAction VisitBuildingBlockStart(BuildingBlock block)
            {
                // Change values by default of created BuildingBlock
                block.Behavior = BuildingBlockBehavior.Paragraph;
                block.Category = "My custom building blocks";
                block.Description =
                    "Using this block in the Quick Parts section of word will place its contents at the cursor.";
                block.Gallery = BuildingBlockGallery.QuickParts;

                block.Guid = Guid.NewGuid();

                // Add content for the BuildingBlock to have an effect when used in the document
                Section section = new Section(mGlossaryDoc);
                block.AppendChild(section);

                Body body = new Body(mGlossaryDoc);
                section.AppendChild(body);

                Paragraph paragraph = new Paragraph(mGlossaryDoc);
                body.AppendChild(paragraph);

                // Add text that will be visible in the document
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

        //ExStart
        //ExFor:BuildingBlocks.GlossaryDocument
        //ExFor:BuildingBlocks.GlossaryDocument.Accept(DocumentVisitor)
        //ExFor:BuildingBlocks.GlossaryDocument.BuildingBlocks
        //ExFor:BuildingBlocks.GlossaryDocument.FirstBuildingBlock
        //ExFor:BuildingBlocks.GlossaryDocument.GetBuildingBlock(BuildingBlocks.BuildingBlockGallery,System.String,System.String)
        //ExFor:BuildingBlocks.GlossaryDocument.LastBuildingBlock
        //ExFor:BuildingBlocks.BuildingBlockCollection
        //ExFor:BuildingBlocks.BuildingBlockCollection.Item(System.Int32)
        //ExFor:BuildingBlocks.BuildingBlockCollection.ToArray
        //ExFor:BuildingBlocks.BuildingBlockGallery
        //ExFor:DocumentVisitor.VisitBuildingBlockEnd(BuildingBlock)
        //ExFor:DocumentVisitor.VisitBuildingBlockStart(BuildingBlock)
        //ExFor:DocumentVisitor.VisitGlossaryDocumentEnd(GlossaryDocument)
        //ExFor:DocumentVisitor.VisitGlossaryDocumentStart(GlossaryDocument)
        //ExSummary:Shows how to use GlossaryDocument and BuildingBlockCollection.
        [Test] //ExSkip
        public void GlossaryDocument()
        {
            Document doc = new Document();

            GlossaryDocument glossaryDoc = new GlossaryDocument();
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 1" });
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 2" });
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 3" });
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 4" });
            glossaryDoc.AppendChild(new BuildingBlock(glossaryDoc) { Name = "Block 5" });

            Assert.AreEqual(5, glossaryDoc.BuildingBlocks.Count);

            doc.GlossaryDocument = glossaryDoc;

            // There is a different ways how to get created building blocks
            Assert.AreEqual("Block 1", glossaryDoc.FirstBuildingBlock.Name);
            Assert.AreEqual("Block 2", glossaryDoc.BuildingBlocks[1].Name);
            Assert.AreEqual("Block 3", glossaryDoc.BuildingBlocks.ToArray()[2].Name);
            Assert.AreEqual("Block 5", glossaryDoc.LastBuildingBlock.Name);

            // Get a block by gallery, category and name
            BuildingBlock block4 =
                glossaryDoc.GetBuildingBlock(BuildingBlockGallery.All, "(Empty Category)", "Block 4");

            // All GUIDs are the same by default
            Assert.AreEqual("00000000-0000-0000-0000-000000000000", block4.Guid.ToString());

            // To be able to uniquely identify blocks by GUID, each GUID must be unique
            // We will do that using a custom visitor
            GlossaryDocVisitor visitor = new GlossaryDocVisitor();
            glossaryDoc.Accept(visitor);

            Assert.AreEqual(5, visitor.GetDictionary().Count);

            Console.WriteLine(visitor.GetText());

            // We can find our new blocks in Microsoft Word via Insert > Quick Parts > Building Blocks Organizer...
            doc.Save(ArtifactsDir + "BuildingBlocks.GlossaryDocument.dotx"); 
        }

        /// <summary>
        /// Simple implementation of giving each building block in a glossary document a unique GUID. Implemented as a Visitor.
        /// </summary>
        public class GlossaryDocVisitor : DocumentVisitor
        {
            public GlossaryDocVisitor()
            {
                mBlocksByGuid = new Dictionary<Guid, BuildingBlock>();
                mBuilder = new StringBuilder();
            }

            public string GetText()
            {
                return mBuilder.ToString();
            }

            public Dictionary<Guid, BuildingBlock> GetDictionary()
            {
                return mBlocksByGuid;
            }

            public override VisitorAction VisitGlossaryDocumentStart(GlossaryDocument glossary)
            {
                mBuilder.AppendLine("Glossary documnent found!");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitGlossaryDocumentEnd(GlossaryDocument glossary)
            {
                mBuilder.AppendLine("Reached end of glossary!");
                mBuilder.AppendLine("BuildingBlocks found: " + mBlocksByGuid.Count);
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBuildingBlockStart(BuildingBlock block)
            {
                block.Guid = Guid.NewGuid();
                mBlocksByGuid.Add(block.Guid, block);
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBuildingBlockEnd(BuildingBlock block)
            {
                mBuilder.AppendLine("\tVisited block \"" + block.Name + "\"");
                mBuilder.AppendLine("\t Type: " + block.Type);
                mBuilder.AppendLine("\t Gallery: " + block.Gallery);
                mBuilder.AppendLine("\t Behavior: " + block.Behavior);
                mBuilder.AppendLine("\t Description: " + block.Description);

                return VisitorAction.Continue;
            }

            private readonly Dictionary<Guid, BuildingBlock> mBlocksByGuid;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd
    }
}