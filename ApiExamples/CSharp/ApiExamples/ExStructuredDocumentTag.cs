// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using NUnit.Framework;
using System.IO;
using System.Linq;
using System.Text;

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

            foreach (StructuredDocumentTag sdTag in sdTags.OfType<StructuredDocumentTag>())
            {
                Console.WriteLine("Type of this SDT is: {0}", sdTag.SdtType);
            }

            //ExEnd
            StructuredDocumentTag sdTagRepeatingSection = (StructuredDocumentTag) sdTags[0];
            Assert.AreEqual(SdtType.RepeatingSection, sdTagRepeatingSection.SdtType);

            StructuredDocumentTag sdTagRichText = (StructuredDocumentTag) sdTags[1];
            Assert.AreEqual(SdtType.RichText, sdTagRichText.SdtType);
        }

        [Test]
        public void SetSpecificStyleToSdt()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.Style
            //ExFor:StructuredDocumentTag.StyleName
            //ExFor:MarkupLevel
            //ExFor:SdtType
            //ExSummary:Shows how to work with styles for content control elements.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Get specific style from the document to apply it to an SDT
            Style quoteStyle = doc.Styles[StyleIdentifier.Quote];
            StructuredDocumentTag sdtPlainText = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
            sdtPlainText.Style = quoteStyle;

            StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Inline);
            sdtRichText.StyleName = "Quote"; // Second method to apply specific style to an SDT control

            // Insert content controls into the document
            builder.InsertNode(sdtPlainText);
            builder.InsertNode(sdtRichText);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            NodeCollection tags = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            foreach (Node node in tags)
            {
                StructuredDocumentTag sdt = (StructuredDocumentTag) node;
                // If style was not defined before, style should be "Default Paragraph Font"
                Assert.AreEqual(StyleIdentifier.Quote, sdt.Style.StyleIdentifier);
                Assert.AreEqual("Quote", sdt.StyleName);
            }
            //ExEnd
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

            StructuredDocumentTag sdt = (StructuredDocumentTag) sdts[0];
            Assert.AreEqual(true, sdt.Checked);
            Assert.That(sdt.XmlMapping.StoreItemId, Is.Empty); //Assert that this sdt has no StoreItemId
        }

        [Test]
        public void ListItemCollection()
        {
            //ExStart
            //ExFor:SdtListItem
            //ExFor:SdtListItem.#ctor(System.String)
            //ExFor:SdtListItem.#ctor(System.String,System.String)
            //ExFor:SdtListItem.DisplayText
            //ExFor:SdtListItem.Value
            //ExFor:SdtListItemCollection
            //ExFor:SdtListItemCollection.Add(Aspose.Words.Markup.SdtListItem)
            //ExFor:SdtListItemCollection.Clear
            //ExFor:SdtListItemCollection.Count
            //ExFor:SdtListItemCollection.GetEnumerator
            //ExFor:SdtListItemCollection.Item(System.Int32)
            //ExFor:SdtListItemCollection.RemoveAt(System.Int32)
            //ExFor:SdtListItemCollection.SelectedValue
            //ExSummary:Shows how to work with StructuredDocumentTag nodes of the DropDownList type.
            // Create a blank document and insert a StructuredDocumentTag that will contain a drop down list
            Document doc = new Document();
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Block);
            doc.FirstSection.Body.AppendChild(tag);

            // A drop down list needs elements, each of which will be a SdtListItem
            SdtListItemCollection listItems = tag.ListItems;
            listItems.Add(new SdtListItem("Value 1"));

            // Each SdtListItem has text that will be displayed when the drop down list is opened, and also a value
            // When we initialize with one string, we are providing just the value
            // If we do that, that value is passed to DisplayText and will consequently be displayed on the screen 
            Assert.AreEqual(listItems[0].DisplayText, listItems[0].Value);

            // Add 3 more SdtListItems with non-empty strings passed to DisplayText
            listItems.Add(new SdtListItem("Item 2", "Value 2"));
            listItems.Add(new SdtListItem("Item 3", "Value 3"));
            listItems.Add(new SdtListItem("Item 4", "Value 4"));

            // We can obtain a count of the SdtListItems and also set the drop down list's SelectedValue attribute to
            // automatically have one of them pre-selected when we open the document in Microsoft Word
            Assert.AreEqual(4, listItems.Count);
            listItems.SelectedValue = listItems[3];

            Assert.AreEqual("Value 4", listItems.SelectedValue.Value);

            // We can enumerate over the collection and print each element
            using (IEnumerator<SdtListItem> enumerator = listItems.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"List item: {enumerator.Current.DisplayText}, value: {enumerator.Current.Value}");
                }
            }

            // We can also remove elements one at a time
            listItems.RemoveAt(3);
            Assert.AreEqual(3, listItems.Count);

            // Make sure to update the SelectedValue's index if it ever ends up out of bounds before saving the document
            listItems.SelectedValue = listItems[1];
           
            doc.Save(ArtifactsDir + "SDT.ListItemCollection.docx");

            // We can clear the whole collection at once too
            listItems.Clear();
            Assert.AreEqual(0, listItems.Count);
            //ExEnd
        }

        [Test]
        [Category("SkipTearDown")]
        public void CreatingCustomXml()
        {
            //ExStart
            //ExFor:CustomXmlPart
            //ExFor:CustomXmlPart.Clone
            //ExFor:CustomXmlPart.Data
            //ExFor:CustomXmlPart.Id
            //ExFor:CustomXmlPart.Schemas
            //ExFor:CustomXmlPartCollection
            //ExFor:CustomXmlPartCollection.Add(CustomXmlPart)
            //ExFor:CustomXmlPartCollection.Add(String, String)
            //ExFor:CustomXmlPartCollection.Clear
            //ExFor:CustomXmlPartCollection.Clone
            //ExFor:CustomXmlPartCollection.Count
            //ExFor:CustomXmlPartCollection.GetById(String)
            //ExFor:CustomXmlPartCollection.GetEnumerator
            //ExFor:CustomXmlPartCollection.Item(Int32)
            //ExFor:CustomXmlPartCollection.RemoveAt(Int32)
            //ExFor:Document.CustomXmlParts
            //ExFor:XmlMapping.SetMapping(CustomXmlPart, String, String)
            //ExSummary:Shows how to create structured document tag with a custom XML data.
            Document doc = new Document();

            // Construct an XML part that contains data and add it to the document's collection
            // Once the "Developer" tab in Mircosoft Word is enabled,
            // we can find elements from this collection as well as a couple defaults in the "XML Mapping Pane" 
            string xmlPartId = Guid.NewGuid().ToString("B");
            string xmlPartContent = "<root><text>Hello, World!</text></root>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlPartContent);

            // The data we entered resides in these variables
            Assert.AreEqual(xmlPartContent.ToCharArray(), xmlPart.Data);
            Assert.AreEqual(xmlPartId, xmlPart.Id);

            // XML parts can be referenced by collection index or GUID
            Assert.AreEqual(xmlPart, doc.CustomXmlParts[0]);
            Assert.AreEqual(xmlPart, doc.CustomXmlParts.GetById(xmlPartId));

            // Once the part is created, we can add XML schema associations like this
            xmlPart.Schemas.Add("http://www.w3.org/2001/XMLSchema");
            
            // We can also clone parts and insert them into the collection directly
            CustomXmlPart xmlPartClone = xmlPart.Clone();
            xmlPartClone.Id = Guid.NewGuid().ToString("B");
            doc.CustomXmlParts.Add(xmlPartClone);

            Assert.AreEqual(2, doc.CustomXmlParts.Count);

            // Iterate through collection with an enumerator and print the contents of each part
            using (IEnumerator<CustomXmlPart> enumerator = doc.CustomXmlParts.GetEnumerator())
            {
                int index = 0;
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"XML part index {index}, ID: {enumerator.Current.Id}");
                    Console.WriteLine($"\tContent: {Encoding.UTF8.GetString(enumerator.Current.Data)}");
                    index++;
                }
            }

            // XML parts can be removed by index
            doc.CustomXmlParts.RemoveAt(1);

            Assert.AreEqual(1, doc.CustomXmlParts.Count);

            // The XML part collection itself can be cloned also
            CustomXmlPartCollection customXmlParts = doc.CustomXmlParts.Clone();

            // And all elements can be cleared like this
            customXmlParts.Clear();

            // Create a StructuredDocumentTag that will display the contents of our part,
            // insert it into the document and save the document
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/text[1]", "");

            doc.FirstSection.Body.AppendChild(sdt);

            doc.Save(ArtifactsDir + "SDT.CustomXml.docx");
            //ExEnd

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "SDT.CustomXml.docx", GoldsDir + "SDT.CustomXml Gold.docx"));
        }

        [Test]
        public void CustomXmlSchemaCollection()
        {
            //ExStart
            //ExFor:CustomXmlSchemaCollection
            //ExFor:CustomXmlSchemaCollection.Add(System.String)
            //ExFor:CustomXmlSchemaCollection.Clear
            //ExFor:CustomXmlSchemaCollection.Clone
            //ExFor:CustomXmlSchemaCollection.Count
            //ExFor:CustomXmlSchemaCollection.GetEnumerator
            //ExFor:CustomXmlSchemaCollection.IndexOf(System.String)
            //ExFor:CustomXmlSchemaCollection.Item(System.Int32)
            //ExFor:CustomXmlSchemaCollection.Remove(System.String)
            //ExFor:CustomXmlSchemaCollection.RemoveAt(System.Int32)
            //ExSummary:Shows how to work with an XML schema collection.
            // Create a document and add a custom XML part
            Document doc = new Document();

            string xmlPartId = Guid.NewGuid().ToString("B");
            string xmlPartContent = "<root><text>Hello, World!</text></root>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlPartContent);

            // Once the part is created, we can add XML schema associations like this,
            // and perform other collection-related operations on the list of schemas for this part
            xmlPart.Schemas.Add("http://www.w3.org/2001/XMLSchema");

            // Collections can be cloned and elements can be added
            CustomXmlSchemaCollection schemas = xmlPart.Schemas.Clone();
            schemas.Add("http://www.w3.org/2001/XMLSchema-instance");
            schemas.Add("http://schemas.microsoft.com/office/2006/metadata/contentType");
            
            Assert.AreEqual(3, schemas.Count);
            Assert.AreEqual(2, schemas.IndexOf(("http://schemas.microsoft.com/office/2006/metadata/contentType")));

            // We can iterate over the collection with an enumerator
            using (IEnumerator<string> enumerator = schemas.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Console.WriteLine(enumerator.Current);
                }
            }

            // We can also remove elements by index, element, or we can clear the entire collection
            schemas.RemoveAt(2);
            schemas.Remove("http://www.w3.org/2001/XMLSchema");
            schemas.Clear();

            Assert.AreEqual(0, schemas.Count);
            //ExEnd
        }

        [Test]
        public void CustomXmlPartStoreItemIdReadOnly()
        {
            //ExStart
            //ExFor:XmlMapping.StoreItemId
            //ExSummary:Shows how to get special id of your xml part.
            Document doc = new Document(ArtifactsDir + "SDT.CustomXml.docx");

            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Console.WriteLine("The Id of your custom xml part is: " + sdt.XmlMapping.StoreItemId);
            //ExEnd
        }

        [Test]
        public void CustomXmlPartStoreItemIdReadOnlyNull()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            sdtCheckBox.Checked = true;

            // Insert content control into the document
            builder.InsertNode(sdtCheckBox);
            
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Console.WriteLine("The Id of your custom xml part is: " + sdt.XmlMapping.StoreItemId);
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

            foreach (StructuredDocumentTag sdt in sdts.OfType<StructuredDocumentTag>())
            {
                sdt.Clear();
            }

            //ExEnd
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            Assert.AreEqual(
                "Enter any content that you want to repeat, including other content controls. You can also insert this control around table rows in order to repeat parts of a table.\r",
                sdts[0].GetText());
            Assert.AreEqual("Click here to enter text.\f", sdts[2].GetText());
        }

        [Test]
        public void SmartTagProperties()
        {
            //ExStart
            //ExFor:CustomXmlProperty.Uri
            //ExFor:CustomXmlPropertyCollection
            //ExFor:CustomXmlPropertyCollection.Add(CustomXmlProperty)
            //ExFor:CustomXmlPropertyCollection.Clear
            //ExFor:CustomXmlPropertyCollection.Contains(String)
            //ExFor:CustomXmlPropertyCollection.Count
            //ExFor:CustomXmlPropertyCollection.GetEnumerator
            //ExFor:CustomXmlPropertyCollection.IndexOfKey(String)
            //ExFor:CustomXmlPropertyCollection.Item(Int32)
            //ExFor:CustomXmlPropertyCollection.Item(String)
            //ExFor:CustomXmlPropertyCollection.Remove(String)
            //ExFor:CustomXmlPropertyCollection.RemoveAt(Int32)
            //ExSummary:Shows how to work with smart tag properties to get in depth information about smart tags.
            // Open a document that contains smart tags and their collection
            Document doc = new Document(MyDir + "SmartTags.doc");

            // Smart tags are an older Microsoft Word feature that can automatically detect and tag
            // any parts of the text that it registers as commonly used information objects such as names, addresses, stock tickers, dates etc
            // In Word 2003, smart tags can be turned on in Tools > AutoCorrect options... > SmartTags tab
            // In our input document there are three objects that were registered as smart tags, but since they can be nested, we have 8 in this collection
            NodeCollection smartTags = doc.GetChildNodes(NodeType.SmartTag, true);
            Assert.AreEqual(8, smartTags.Count);

            // The last smart tag is of the "Date" type, which we will retrieve here
            SmartTag smartTag = (SmartTag)smartTags[7];

            // The Properties attribute, for some smart tags, elaborates on the text object that Word picked up as a smart tag
            // In the case of our "Date" smart tag, its properties will let us know the year, month and day within the smart tag
            CustomXmlPropertyCollection properties = smartTag.Properties;

            // We can enumerate over the collection and print the aforementioned properties to the console
            Assert.AreEqual(4, properties.Count);

            using (IEnumerator<CustomXmlProperty> enumerator = properties.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"Property name: {enumerator.Current.Name}, value: {enumerator.Current.Value}");
                    Assert.AreEqual("", enumerator.Current.Uri);
                }
            }
            
            // We can also access the elements in various ways, including as a key-value pair
            Assert.True(properties.Contains("Day"));
            Assert.AreEqual("22", properties["Day"].Value);
            Assert.AreEqual("2003", properties[2].Value);
            Assert.AreEqual(1, properties.IndexOfKey("Month"));

            // We can also remove elements by name, index or clear the collection entirely
            properties.RemoveAt(3);
            properties.Remove("Year");
            Assert.AreEqual(2, (properties.Count));

            properties.Clear();
            Assert.AreEqual(0, (properties.Count));

            // We can remove the entire smart tag like this
            smartTag.Remove();
            //ExEnd
        }

        //ExStart
        //ExFor:CustomXmlProperty
        //ExFor:CustomXmlProperty.#ctor(String,String,String)
        //ExFor:CustomXmlProperty.Name
        //ExFor:CustomXmlProperty.Value
        //ExFor:Markup.SmartTag
        //ExFor:Markup.SmartTag.#ctor(Aspose.Words.DocumentBase)
        //ExFor:Markup.SmartTag.Accept(Aspose.Words.DocumentVisitor)
        //ExFor:Markup.SmartTag.Element
        //ExFor:Markup.SmartTag.Properties
        //ExFor:Markup.SmartTag.Uri
        //ExSummary:Shows how to create smart tags.
        [Test] //ExSkip
        public void SmartTags()
        {
            Document doc = new Document();
            SmartTag smartTag = new SmartTag(doc);
            smartTag.Element = "date";

            // Specify a date and set smart tag properties accordingly
            smartTag.AppendChild(new Run(doc, "May 29, 2019"));

            smartTag.Properties.Add(new CustomXmlProperty("Day", "", "29"));
            smartTag.Properties.Add(new CustomXmlProperty("Month", "", "5"));
            smartTag.Properties.Add(new CustomXmlProperty("Year", "", "2019"));

            // Set the smart tag's uri to the default
            smartTag.Uri = "urn:schemas-microsoft-com:office:smarttags";

            doc.FirstSection.Body.FirstParagraph.AppendChild(smartTag);
            doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, " is a date. "));

            // Create and add one more smart tag, this time for a financial symbol
            smartTag = new SmartTag(doc);
            smartTag.Element = "stockticker";
            smartTag.Uri = "urn:schemas-microsoft-com:office:smarttags";

            smartTag.AppendChild(new Run(doc, "MSFT"));

            doc.FirstSection.Body.FirstParagraph.AppendChild(smartTag);
            doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, " is a stock ticker."));

            // Print all the smart tags in our document with a document visitor
            doc.Accept(new SmartTagVisitor());

            doc.Save(ArtifactsDir + "SmartTags.doc");
            //ExEnd
        }

        /// <summary>
        /// DocumentVisitor implementation that prints smart tags and their contents
        /// </summary>
        private class SmartTagVisitor : DocumentVisitor
        {
            /// <summary>
            /// Called when a SmartTag node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSmartTagStart(SmartTag smartTag)
            {
                Console.WriteLine($"Smart tag type: {smartTag.Element}");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a SmartTag node is ended.
            /// </summary>
            public override VisitorAction VisitSmartTagEnd(SmartTag smartTag)
            {
                Console.WriteLine($"\tContents: \"{smartTag.ToString(SaveFormat.Text)}\"");

                if (smartTag.Properties.Count == 0)
                {
                    Console.WriteLine("\tContains no properties");

                }
                else
                {
                    Console.Write("\tProperties: ");
                    string[] properties = new string[smartTag.Properties.Count];
                    int index = 0;         
                    
                    foreach (CustomXmlProperty cxp in smartTag.Properties)
                    {
                        properties[index++] = $"\"{cxp.Name}\" = \"{cxp.Value}\"";
                    }

                    Console.WriteLine(String.Join(", ", properties));
                }

                return VisitorAction.Continue;
            }
        }
        //ExEnd

        [Test]
        public void AccessToBuildingBlockPropertiesFromDocPartObjSdt()
        {
            Document doc = new Document(MyDir + "StructuredDocumentTag.BuildingBlocks.docx");

            StructuredDocumentTag docPartObjSdt =
                (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            Assert.AreEqual(SdtType.DocPartObj, docPartObjSdt.SdtType);
            Assert.AreEqual("Table of Contents", docPartObjSdt.BuildingBlockGallery);
        }

        [Test]
        public void AccessToBuildingBlockPropertiesFromPlainTextSdt()
        {
            Document doc = new Document(MyDir + "StructuredDocumentTag.BuildingBlocks.docx");

            StructuredDocumentTag plainTextSdt =
                (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 1, true);

            Assert.AreEqual(SdtType.PlainText, plainTextSdt.SdtType);
            Assert.That(() => plainTextSdt.BuildingBlockGallery, Throws.TypeOf<InvalidOperationException>(),
                "BuildingBlockType is only accessible for BuildingBlockGallery SDT type.");
        }

        [Test]
        public void AccessToBuildingBlockPropertiesFromBuildingBlockGallerySdtType()
        {
            Document doc = new Document();

            StructuredDocumentTag buildingBlockSdt =
                new StructuredDocumentTag(doc, SdtType.BuildingBlockGallery, MarkupLevel.Block)
                {
                    BuildingBlockCategory = "Built-in",
                    BuildingBlockGallery = "Table of Contents"
                };

            doc.FirstSection.Body.AppendChild(buildingBlockSdt);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            buildingBlockSdt =
                (StructuredDocumentTag) doc.FirstSection.Body.GetChild(NodeType.StructuredDocumentTag, 0, true);

            Assert.AreEqual(SdtType.BuildingBlockGallery, buildingBlockSdt.SdtType);
            Assert.AreEqual("Table of Contents", buildingBlockSdt.BuildingBlockGallery);
            Assert.AreEqual("Built-in", buildingBlockSdt.BuildingBlockCategory);
        }
    }
}