using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExBuildingBlocks : ApiExampleBase
    {
        [Test]
        public void BuildingBlocks()
        {
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
            //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockCollection
            //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockCollection.Item(System.Int32)
            //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockCollection.ToArray
            //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockGallery
            //ExFor:Aspose.Words.BuildingBlocks.BuildingBlockType
            //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument
            //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.Accept(Aspose.Words.DocumentVisitor)
            //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.BuildingBlocks
            //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.FirstBuildingBlock
            //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.GetBuildingBlock(Aspose.Words.BuildingBlocks.BuildingBlockGallery,System.String,System.String)
            //ExFor:Aspose.Words.BuildingBlocks.GlossaryDocument.LastBuildingBlock
            //ExFor:Aspose.Words.BuildingBlocks.NamespaceDoc
            //ExStart
            Document doc = new Document();

            GlossaryDocument glossaryDocument = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDocument;

            BuildingBlock buildingBlock = new BuildingBlock(glossaryDocument);
            glossaryDocument.AppendChild(buildingBlock);

            Assert.AreEqual("(Empty Name)", buildingBlock.Name);
            Assert.AreEqual("(Empty Category)", buildingBlock.Category);
            Assert.AreEqual(BuildingBlockBehavior.Content, buildingBlock.Behavior);

            buildingBlock.Behavior = BuildingBlockBehavior.Page;
            buildingBlock.Description = "Building block description";
            buildingBlock.Name = "MyBBName";
            buildingBlock.Gallery = BuildingBlockGallery.QuickParts;
            buildingBlock.Category = "General";
            buildingBlock.Type = BuildingBlockType.AutoCorrect;
            //ExEnd
        }
    }
}
