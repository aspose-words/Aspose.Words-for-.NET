// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using Aspose.Words.Markup;
using NUnit.Framework;
using System.IO;

namespace ApiExamples
{
    /// <summary>
    /// Tests that verify work with structured document tags in the document 
    /// </summary>
    [TestFixture]
    internal class ExStructuredDocumentTag : ApiExampleBase
    {
        [Test]
        public void RepeatingSection()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.SdtType
            //ExSummary:Shows how to get type of structured document tag.
            Document doc = new Document(MyDir + "TestRepeatingSection.docx");

            NodeCollection sdTags = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            foreach (StructuredDocumentTag sdTag in sdTags)
            {
                Console.WriteLine("Type of this SDT is: {0}",sdTag.SdtType);
            }
            //ExEnd
            StructuredDocumentTag sdTagRepeatingSection = (StructuredDocumentTag)sdTags[0];
            Assert.AreEqual(SdtType.RepeatingSection, sdTagRepeatingSection.SdtType);

            StructuredDocumentTag sdTagRichText = (StructuredDocumentTag)sdTags[1];
            Assert.AreEqual(SdtType.RichText, sdTagRichText.SdtType);
        }

        [Test]
        public void CheckBox()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.#ctor(DocumentBase, SdtType, MarkupLevel)
            //ExFor:StructuredDocumentTag.Checked
            //ExSummary:Show how to create and insert checkbox structured document tag.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            sdtCheckBox.Checked = true;

            // Insert content control into the document
            builder.InsertNode(sdtCheckBox);
            //ExEnd
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            NodeCollection sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            StructuredDocumentTag sdt = (StructuredDocumentTag)sdts[0];
            Assert.AreEqual(true, sdt.Checked);
        }

        [Test]
        public void CreatingCustomXml()
        {
            //ExStart
            //ExFor:CustomXmlPart
            //ExFor:CustomXmlPartCollection.Add(String, String)
            //ExFor:Document.CustomXmlParts
            //ExFor:XmlMapping.SetMapping(CustomXmlPart, String, String)
            //ExSummary:Shows how to create structured document tag with a custom XML data.
            Document doc = new Document();
            // Add test XML data part to the collection.
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), "<root><text>Hello, World!</text></root>");

            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/text[1]", "");

            doc.FirstSection.Body.AppendChild(sdt);

            doc.Save(MyDir + @"\Artifacts\SDT.CustomXml.docx");
            //ExEnd
            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\SDT.CustomXml.docx", MyDir + @"\Golds\SDT.CustomXml Gold.docx"));
        }

        [Test]
        public void ClearTextFromStructuredDocumentTags()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.Clear
            //ExSummary:Shows how to delete content of StructuredDocumentTag elements.
            Document doc = new Document(MyDir + "TestRepeatingSection.docx");

            NodeCollection sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            Assert.IsNotNull(sdts);

            foreach (StructuredDocumentTag sdt in sdts)
            {
                sdt.Clear();
            }
            //ExEnd
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            Assert.AreEqual("Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.\r", sdts[0].GetText());
            Assert.AreEqual("Click here to enter text.\f", sdts[2].GetText());
        }

        [Test]
        public void AccessToBuildingBlockPropertiesFromDocPartObjSdt()
        {
            Document doc = new Document(MyDir + "StructuredDocumentTag.BuildingBlocks.docx");

            StructuredDocumentTag docPartObjSdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            Assert.AreEqual(SdtType.DocPartObj, docPartObjSdt.SdtType);
            Assert.AreEqual("Table of Contents", docPartObjSdt.BuildingBlockGallery);
        }

        [Test]
        public void AccessToBuildingBlockPropertiesFromPlainTextSdt()
        {
            Document doc = new Document(MyDir + "StructuredDocumentTag.BuildingBlocks.docx");

            StructuredDocumentTag plainTextSdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 1, true);

            Assert.AreEqual(SdtType.PlainText, plainTextSdt.SdtType);
            Assert.That(() => plainTextSdt.BuildingBlockGallery, Throws.TypeOf<InvalidOperationException>(), "BuildingBlockType is only accessible for BuildingBlockGallery SDT type.");
        }

        [Test]
        public void AccessToBuildingBlockPropertiesFromBuildingBlockGallerySdtType()
        {
            Document doc = new Document();

            StructuredDocumentTag buildingBlockSdt = new StructuredDocumentTag(doc, SdtType.BuildingBlockGallery, MarkupLevel.Block)
            {
                BuildingBlockCategory = "Built-in",
                BuildingBlockGallery = "Table of Contents"
            };
            
            doc.FirstSection.Body.AppendChild(buildingBlockSdt);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            buildingBlockSdt = (StructuredDocumentTag)doc.FirstSection.Body.GetChild(NodeType.StructuredDocumentTag, 0, true);

            Assert.AreEqual(SdtType.BuildingBlockGallery, buildingBlockSdt.SdtType);
            Assert.AreEqual("Table of Contents", buildingBlockSdt.BuildingBlockGallery);
            Assert.AreEqual("Built-in", buildingBlockSdt.BuildingBlockCategory);
        }
    }
}
