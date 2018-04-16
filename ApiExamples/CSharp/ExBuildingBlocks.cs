using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExBuildingBlocks : ApiExampleBase
    {
        [Test]
        public void BuildingBlocks()
        {
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
            Document doc = new Document();

            // BuildingBlocks live inside the glossary document.
            // If you're making a document from scratch, the glossary document must also be manually created.
            GlossaryDocument glossaryDoc = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDoc;

            // Create a building block and edit attributes.
            BuildingBlock block = new BuildingBlock(glossaryDoc);

            block.Name = "Custom Block 1";
            block.Behavior = BuildingBlockBehavior.Paragraph;
            block.Category = "My building blocks";
            block.Description = "Using this block in the Quick Parts section of word will place its contents at the cursor.";
            block.Gallery = BuildingBlockGallery.QuickParts;
            Assert.AreEqual("00000000-0000-0000-0000-000000000000", block.Guid.ToString());
            Assert.AreEqual(BuildingBlockType.None, block.Type);

            glossaryDoc.AppendChild(block);

            // Add content for the block for it to have an effect when used in the document.
            Section section = new Section(glossaryDoc);
            block.AppendChild(section);

            Body body = new Body(glossaryDoc);
            section.AppendChild(body);

            Paragraph paragraph = new Paragraph(glossaryDoc);
            body.AppendChild(paragraph);

            Run run = new Run(glossaryDoc, "Text added by Block 1!");
            paragraph.AppendChild(run);

            doc.Save(MyDir + @"BuildingBlocks.dotx");
            //ExEnd
        }

        [Test]
        public void GlossaryDocument()
        {
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
            //ExSummary: Shows how to use GlossaryDocument and BuildingBlockCollection.
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
            Assert.AreEqual("00000000-0000-0000-0000-000000000000", glossaryDoc.GetBuildingBlock(BuildingBlockGallery.All, "(Empty Category)", "Block 4").Guid.ToString());
            Assert.AreEqual("Block 5", glossaryDoc.LastBuildingBlock.Name);
            //ExEnd
        }
    }
}
