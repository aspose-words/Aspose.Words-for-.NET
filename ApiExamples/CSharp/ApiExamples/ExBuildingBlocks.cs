// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
        public void CreateAndInsert()
        {
            // A document's glossary document stores building blocks.
            Document doc = new Document();
            GlossaryDocument glossaryDoc = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDoc;

            // Create a building block, name it, and then add it to the glossary document.
            BuildingBlock block = new BuildingBlock(glossaryDoc)
            {
                Name = "Custom Block"
            };

            glossaryDoc.AppendChild(block);

            // All new building block GUIDs have the same zero value by default, and we can give them a new unique value.
            Assert.AreEqual("00000000-0000-0000-0000-000000000000", block.Guid.ToString());

            block.Guid = Guid.NewGuid();

            // The following properties categorize building blocks
            // in the menu we can access in Microsoft Word via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
            Assert.AreEqual("(Empty Category)", block.Category);
            Assert.AreEqual(BuildingBlockType.None, block.Type);
            Assert.AreEqual(BuildingBlockGallery.All, block.Gallery);
            Assert.AreEqual(BuildingBlockBehavior.Content, block.Behavior);

            // Before we can add this building block to our document, we will need to give it some contents,
            // which we will do using a document visitor. This visitor will also set a category, gallery, and behavior.
            BuildingBlockVisitor visitor = new BuildingBlockVisitor(glossaryDoc);
            block.Accept(visitor);

            // We can access the block that we just made from the glossary document.
            BuildingBlock customBlock = glossaryDoc.GetBuildingBlock(BuildingBlockGallery.QuickParts,
                "My custom building blocks", "Custom Block");

            // The block itself is a section that contains the text.
            Assert.AreEqual($"Text inside {customBlock.Name}\f", customBlock.FirstSection.Body.FirstParagraph.GetText());
            Assert.AreEqual(customBlock.FirstSection, customBlock.LastSection);
            Assert.DoesNotThrow(() => Guid.Parse(customBlock.Guid.ToString())); //ExSkip
            Assert.AreEqual("My custom building blocks", customBlock.Category); //ExSkip
            Assert.AreEqual(BuildingBlockType.None, customBlock.Type); //ExSkip
            Assert.AreEqual(BuildingBlockGallery.QuickParts, customBlock.Gallery); //ExSkip
            Assert.AreEqual(BuildingBlockBehavior.Paragraph, customBlock.Behavior); //ExSkip

            // Now, we can insert it into the document as a new section.
            doc.AppendChild(doc.ImportNode(customBlock.FirstSection, true));

            // We can also find it in Microsoft Word's Building Blocks Organizer and place it manually.
            doc.Save(ArtifactsDir + "BuildingBlocks.CreateAndInsert.dotx");
        }

        /// <summary>
        /// Sets up a visited building block to be inserted into the document as a quick part and adds text to its contents.
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
                // Configure the building block as a quick part, and add properties used by Building Blocks Organizer.
                block.Behavior = BuildingBlockBehavior.Paragraph;
                block.Category = "My custom building blocks";
                block.Description =
                    "Using this block in the Quick Parts section of word will place its contents at the cursor.";
                block.Gallery = BuildingBlockGallery.QuickParts;

                // Add a section with text.
                // Inserting the block into the document will append this section with its child nodes at the location.
                Section section = new Section(mGlossaryDoc);
                block.AppendChild(section);
                block.FirstSection.EnsureMinimum();

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
        //ExSummary:Shows ways of accessing building blocks in a glossary document.
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

            // There are various ways of accessing building blocks.
            // 1 -  Get the first/last building blocks in the collection:
            Assert.AreEqual("Block 1", glossaryDoc.FirstBuildingBlock.Name);
            Assert.AreEqual("Block 5", glossaryDoc.LastBuildingBlock.Name);

            // 2 -  Get a building block by index:
            Assert.AreEqual("Block 2", glossaryDoc.BuildingBlocks[1].Name);
            Assert.AreEqual("Block 3", glossaryDoc.BuildingBlocks.ToArray()[2].Name);

            // 3 -  Get the first building block that matches a gallery, name and category:
            Assert.AreEqual("Block 4", 
                glossaryDoc.GetBuildingBlock(BuildingBlockGallery.All, "(Empty Category)", "Block 4").Name);

            // We will do that using a custom visitor,
            // which will give every BuildingBlock in the GlossaryDocument a unique GUID
            GlossaryDocVisitor visitor = new GlossaryDocVisitor();
            glossaryDoc.Accept(visitor);
            Assert.AreEqual(5, visitor.GetDictionary().Count); //ExSkip

            Console.WriteLine(visitor.GetText());

            // In Microsoft Word, we can access the building blocks via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
            doc.Save(ArtifactsDir + "BuildingBlocks.GlossaryDocument.dotx"); 
        }

        /// <summary>
        /// Gives each building block in a visited glossary document a unique GUID.
        /// Stores the GUID-building block pairs in a dictionary.
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
                mBuilder.AppendLine("Glossary document found!");
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
                Assert.AreEqual("00000000-0000-0000-0000-000000000000", block.Guid.ToString()); //ExSkip
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