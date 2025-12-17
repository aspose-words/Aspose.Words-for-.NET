// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
        //ExFor:BuildingBlock
        //ExFor:BuildingBlock.#ctor(GlossaryDocument)
        //ExFor:BuildingBlock.Accept(DocumentVisitor)
        //ExFor:BuildingBlock.AcceptStart(DocumentVisitor)
        //ExFor:BuildingBlock.AcceptEnd(DocumentVisitor)
        //ExFor:BuildingBlock.Behavior
        //ExFor:BuildingBlock.Category
        //ExFor:BuildingBlock.Description
        //ExFor:BuildingBlock.FirstSection
        //ExFor:BuildingBlock.Gallery
        //ExFor:BuildingBlock.Guid
        //ExFor:BuildingBlock.LastSection
        //ExFor:BuildingBlock.Name
        //ExFor:BuildingBlock.Sections
        //ExFor:BuildingBlock.Type
        //ExFor:BuildingBlockBehavior
        //ExFor:BuildingBlockType
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
            Assert.That(block.Guid.ToString(), Is.EqualTo("00000000-0000-0000-0000-000000000000"));

            block.Guid = Guid.NewGuid();

            // The following properties categorize building blocks
            // in the menu we can access in Microsoft Word via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
            Assert.That(block.Category, Is.EqualTo("(Empty Category)"));
            Assert.That(block.Type, Is.EqualTo(BuildingBlockType.None));
            Assert.That(block.Gallery, Is.EqualTo(BuildingBlockGallery.All));
            Assert.That(block.Behavior, Is.EqualTo(BuildingBlockBehavior.Content));

            // Before we can add this building block to our document, we will need to give it some contents,
            // which we will do using a document visitor. This visitor will also set a category, gallery, and behavior.
            BuildingBlockVisitor visitor = new BuildingBlockVisitor(glossaryDoc);
            // Visit start/end of the BuildingBlock.
            block.Accept(visitor);

            // We can access the block that we just made from the glossary document.
            BuildingBlock customBlock = glossaryDoc.GetBuildingBlock(BuildingBlockGallery.QuickParts,
                "My custom building blocks", "Custom Block");

            // The block itself is a section that contains the text.
            Assert.That(customBlock.FirstSection.Body.FirstParagraph.GetText(), Is.EqualTo($"Text inside {customBlock.Name}\f"));
            Assert.That(customBlock.LastSection, Is.EqualTo(customBlock.FirstSection));
            Assert.DoesNotThrow(() => Guid.Parse(customBlock.Guid.ToString())); //ExSkip
            Assert.That(customBlock.Category, Is.EqualTo("My custom building blocks")); //ExSkip
            Assert.That(customBlock.Type, Is.EqualTo(BuildingBlockType.None)); //ExSkip
            Assert.That(customBlock.Gallery, Is.EqualTo(BuildingBlockGallery.QuickParts)); //ExSkip
            Assert.That(customBlock.Behavior, Is.EqualTo(BuildingBlockBehavior.Paragraph)); //ExSkip

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
        //ExFor:GlossaryDocument
        //ExFor:GlossaryDocument.Accept(DocumentVisitor)
        //ExFor:GlossaryDocument.AcceptStart(DocumentVisitor)
        //ExFor:GlossaryDocument.AcceptEnd(DocumentVisitor)
        //ExFor:GlossaryDocument.BuildingBlocks
        //ExFor:GlossaryDocument.FirstBuildingBlock
        //ExFor:GlossaryDocument.GetBuildingBlock(BuildingBlockGallery,String,String)
        //ExFor:GlossaryDocument.LastBuildingBlock
        //ExFor:BuildingBlockCollection
        //ExFor:BuildingBlockCollection.Item(Int32)
        //ExFor:BuildingBlockCollection.ToArray
        //ExFor:BuildingBlockGallery
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

            BuildingBlock child1 = new BuildingBlock(glossaryDoc) { Name = "Block 1" };
            glossaryDoc.AppendChild(child1);
            BuildingBlock child2 = new BuildingBlock(glossaryDoc) { Name = "Block 2" };
            glossaryDoc.AppendChild(child2);
            BuildingBlock child3 = new BuildingBlock(glossaryDoc) { Name = "Block 3" };
            glossaryDoc.AppendChild(child3);
            BuildingBlock child4 = new BuildingBlock(glossaryDoc) { Name = "Block 4" };
            glossaryDoc.AppendChild(child4);
            BuildingBlock child5 = new BuildingBlock(glossaryDoc) { Name = "Block 5" };
            glossaryDoc.AppendChild(child5);

            Assert.That(glossaryDoc.BuildingBlocks.Count, Is.EqualTo(5));

            doc.GlossaryDocument = glossaryDoc;

            // There are various ways of accessing building blocks.
            // 1 -  Get the first/last building blocks in the collection:
            Assert.That(glossaryDoc.FirstBuildingBlock.Name, Is.EqualTo("Block 1"));
            Assert.That(glossaryDoc.LastBuildingBlock.Name, Is.EqualTo("Block 5"));

            // 2 -  Get a building block by index:
            Assert.That(glossaryDoc.BuildingBlocks[1].Name, Is.EqualTo("Block 2"));
            Assert.That(glossaryDoc.BuildingBlocks.ToArray()[2].Name, Is.EqualTo("Block 3"));

            // 3 -  Get the first building block that matches a gallery, name and category:
            Assert.That(glossaryDoc.GetBuildingBlock(BuildingBlockGallery.All, "(Empty Category)", "Block 4").Name, Is.EqualTo("Block 4"));

            // We will do that using a custom visitor,
            // which will give every BuildingBlock in the GlossaryDocument a unique GUID
            GlossaryDocVisitor visitor = new GlossaryDocVisitor();
            // Visit start/end of the Glossary document.
            glossaryDoc.Accept(visitor);
            // Visit only start of the Glossary document.
            glossaryDoc.AcceptStart(visitor);
            // Visit only end of the Glossary document.
            glossaryDoc.AcceptEnd(visitor);
            Assert.That(visitor.GetDictionary().Count, Is.EqualTo(5)); //ExSkip

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
                Assert.That(block.Guid.ToString(), Is.EqualTo("00000000-0000-0000-0000-000000000000")); //ExSkip
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