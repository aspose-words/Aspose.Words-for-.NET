// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Markup;
using NUnit.Framework;
using System.Linq;
using System.Text;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
#if NET462 || NETCOREAPP2_1 || JAVA
using Aspose.Pdf.Text;
#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExStructuredDocumentTag : ApiExampleBase
    {
        [Test]
        public void RepeatingSection()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.SdtType
            //ExSummary:Shows how to get the type of a structured document tag.
            Document doc = new Document(MyDir + "Structured document tags.docx");

            List<StructuredDocumentTag> sdTags = doc.GetChildNodes(NodeType.StructuredDocumentTag, true).OfType<StructuredDocumentTag>().ToList();

            Assert.AreEqual(SdtType.RepeatingSection, sdTags[0].SdtType);
            Assert.AreEqual(SdtType.RepeatingSectionItem, sdTags[1].SdtType);
            Assert.AreEqual(SdtType.RichText, sdTags[2].SdtType);
            //ExEnd
        }

        [Test]
        public void ApplyStyle()
        {
            //ExStart
            //ExFor:StructuredDocumentTag
            //ExFor:StructuredDocumentTag.NodeType
            //ExFor:StructuredDocumentTag.Style
            //ExFor:StructuredDocumentTag.StyleName
            //ExFor:MarkupLevel
            //ExFor:SdtType
            //ExSummary:Shows how to work with styles for content control elements.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Below are two ways tp apply a style from the document to a structured document tag.
            // 1 -  Apply a style object from the document's style collection:
            Style quoteStyle = doc.Styles[StyleIdentifier.Quote];
            StructuredDocumentTag sdtPlainText = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
            sdtPlainText.Style = quoteStyle;

            // 2 -  Reference a style in the document by name:
            StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Inline);
            sdtRichText.StyleName = "Quote";

            builder.InsertNode(sdtPlainText);
            builder.InsertNode(sdtRichText);

            Assert.AreEqual(NodeType.StructuredDocumentTag, sdtPlainText.NodeType);

            NodeCollection tags = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            foreach (Node node in tags)
            {
                StructuredDocumentTag sdt = (StructuredDocumentTag)node;

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
            //ExSummary:Show how to create a structured document tag in the form of a check box.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            sdtCheckBox.Checked = true;

            builder.InsertNode(sdtCheckBox);

            doc.Save(ArtifactsDir + "StructuredDocumentTag.CheckBox.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "StructuredDocumentTag.CheckBox.docx");

            StructuredDocumentTag[] sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true).OfType<StructuredDocumentTag>().ToArray();

            Assert.AreEqual(true, sdts[0].Checked);
            Assert.That(sdts[0].XmlMapping.StoreItemId, Is.Empty);
        }

#if NET462 || NETCOREAPP2_1 || JAVA // because of a Xamarin bug with CultureInfo (https://xamarin.github.io/bugzilla-archives/59/59077/bug.html)
        [Test, Category("SkipMono")]
        public void Date()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.CalendarType
            //ExFor:StructuredDocumentTag.DateDisplayFormat
            //ExFor:StructuredDocumentTag.DateDisplayLocale
            //ExFor:StructuredDocumentTag.DateStorageFormat
            //ExFor:StructuredDocumentTag.FullDate
            //ExSummary:Shows how to prompt the user to enter a date with a StructuredDocumentTag.
            Document doc = new Document();

            // Insert a StructuredDocumentTag that prompts the user to enter a date.
            // In Microsoft Word, this element is known as a "Date picker content control".
            // When we click on the arrow on the right end of this tag in Microsoft Word,
            // we will see a pop up in the form of a clickable calendar.
            // We can use that popup to select a date that the tag will display.
            StructuredDocumentTag sdtDate = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline);

            // Display the date according to the Saudi Arabian Arabic locale.
            sdtDate.DateDisplayLocale = CultureInfo.GetCultureInfo("ar-SA").LCID;

            // Set the format with which to display the date.
            sdtDate.DateDisplayFormat = "dd MMMM, yyyy";
            sdtDate.DateStorageFormat = SdtDateStorageFormat.DateTime;

            // Display the date according to the Hijri calendar.
            sdtDate.CalendarType = SdtCalendarType.Hijri;

            // Before the user chooses a date in Microsoft Word, the tag will display the text "Click here to enter a date.".
            // Set the "FullDate" property to get the tag to display a default date, according to the tag's calendar.
            sdtDate.FullDate = new DateTime(1440, 10, 20);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertNode(sdtDate);

            doc.Save(ArtifactsDir + "StructuredDocumentTag.Date.docx");
            //ExEnd
        }
#endif

        [Test]
        public void PlainText()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.Color
            //ExFor:StructuredDocumentTag.ContentsFont
            //ExFor:StructuredDocumentTag.EndCharacterFont
            //ExFor:StructuredDocumentTag.Id
            //ExFor:StructuredDocumentTag.Level
            //ExFor:StructuredDocumentTag.Multiline
            //ExFor:StructuredDocumentTag.Tag
            //ExFor:StructuredDocumentTag.Title
            //ExFor:StructuredDocumentTag.RemoveSelfOnly
            //ExSummary:Shows how to create a structured document tag in the form of a plain text box and modify its appearance.
            Document doc = new Document();

            // Create a structured document tag that will contain plain text.
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

            // Set the title and color of the frame that appears when
            // you mouse over the structured document tag in Microsoft Word.
            tag.Title = "My plain text";
            tag.Color = Color.Magenta;

            // Set a tag for this structured document tag, which is obtainable
            // as an XML element named "tag", with the string below in its "@val" attribute.
            tag.Tag = "MyPlainTextSDT";

            // Every structured document tag has a random unique ID.
            Assert.That(tag.Id, Is.Positive);

            // Set the font for the text inside the structured document tag.
            tag.ContentsFont.Name = "Arial";

            // Set the font for the text at the end of the structured document tag.
            // Any text that we type in the document body after moving out of the tag with arrow keys will use this font.
            tag.EndCharacterFont.Name = "Arial Black";

            // By default, this is false and pressing enter while inside a structured document tag does nothing
            // When set to true, our StructuredDocumentTag can have multiple lines

            // Set the "Multiline" property to "false" to only allow the contents
            // of this structured document tag to span a single line.
            // Set the "Multiline" property to "true" to allow the tag to contain multiple lines of content.
            tag.Multiline = true;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertNode(tag);

            // Insert a clone of our structured document tag in a new paragraph.
            StructuredDocumentTag tagClone = (StructuredDocumentTag)tag.Clone(true);
            builder.InsertParagraph();
            builder.InsertNode(tagClone);

            // Use the "RemoveSelfOnly" method to remove a structured document tag, while keeping its contents in the document.
            tagClone.RemoveSelfOnly();

            doc.Save(ArtifactsDir + "StructuredDocumentTag.PlainText.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "StructuredDocumentTag.PlainText.docx");
            tag = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            Assert.AreEqual("My plain text", tag.Title);
            Assert.AreEqual(Color.Magenta.ToArgb(), tag.Color.ToArgb());
            Assert.AreEqual("MyPlainTextSDT", tag.Tag);
            Assert.That(tag.Id, Is.Positive);
            Assert.AreEqual("Arial", tag.ContentsFont.Name);
            Assert.AreEqual("Arial Black", tag.EndCharacterFont.Name);
            Assert.True(tag.Multiline);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void IsTemporary(bool isTemporary)
        {
            //ExStart
            //ExFor:StructuredDocumentTag.IsTemporary
            //ExSummary:Shows how to make single-use controls.
            Document doc = new Document();

            // Insert a plain text StructuredDocumentTag,
            // which will act as a plain text form that the user may enter text into.
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

            // Set the "IsTemporary" property to "true" to make the structured document tag disappear and
            // assimilate its contents into the document after the user edits it once in Microsoft Word.
            // Set the "IsTemporary" property to "false" to allow the user to edit the contents
            // of the structured document tag any number of times.
            tag.IsTemporary = isTemporary;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Please enter text: ");
            builder.InsertNode(tag);

            // Insert another structured document tag in the form of a check box, and set its default state to "checked".
            tag = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            tag.Checked = true;

            // Set the "IsTemporary" property to "true" to make the check box become a symbol
            // once the user clicks on it in Microsoft Word.
            // Set the "IsTemporary" property to "false" to allow the user to click on the check box any number of times.
            tag.IsTemporary = isTemporary;

            builder.Write("\nPlease click the check box: ");
            builder.InsertNode(tag);

            doc.Save(ArtifactsDir + "StructuredDocumentTag.IsTemporary.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "StructuredDocumentTag.IsTemporary.docx");

            Assert.AreEqual(2, 
                doc.GetChildNodes(NodeType.StructuredDocumentTag, true).Count(sdt => ((StructuredDocumentTag)sdt).IsTemporary == isTemporary));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void PlaceholderBuildingBlock(bool isShowingPlaceholderText)
        {
            //ExStart
            //ExFor:StructuredDocumentTag.IsShowingPlaceholderText
            //ExFor:StructuredDocumentTag.Placeholder
            //ExFor:StructuredDocumentTag.PlaceholderName
            //ExSummary:Shows how to use the contents of a building block as a custom placeholder text for a structured document tag. 
            Document doc = new Document();

            // Insert a plain text structured document tag of the "PlainText" type, which will function like a text box.
            // The contents that it will display by default are a "Click here to enter text." prompt.
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

            // We can get the tag to display the contents of a building block instead of the default text.
            // First, add a building block with contents to the glossary document.
            GlossaryDocument glossaryDoc = doc.GlossaryDocument;

            BuildingBlock substituteBlock = new BuildingBlock(glossaryDoc);
            substituteBlock.Name = "Custom Placeholder";
            substituteBlock.AppendChild(new Section(glossaryDoc));
            substituteBlock.FirstSection.AppendChild(new Body(glossaryDoc));
            substituteBlock.FirstSection.Body.AppendParagraph("Custom placeholder text.");

            glossaryDoc.AppendChild(substituteBlock);

            // Then, use the structured document tag's "PlaceholderName" property to reference that building block by name.
            tag.PlaceholderName = "Custom Placeholder";

            // If "PlaceholderName" refers to an existing block in the parent document's glossary document,
            // we will be able to verify the building block via the "Placeholder" property.
            Assert.AreEqual(substituteBlock, tag.Placeholder);

            // Set the "IsShowingPlaceholderText" property to "true" to treat the
            // current contents of the structured document tag as placeholder text.
            // This means that clicking on the text box in Microsoft Word will immediately highlight all of the tag's contents.
            // Set the "IsShowingPlaceholderText" property to "false" to get the
            // structured document tag to treat its contents as text that a user has already entered.
            // Clicking on this text in Microsoft Word will place the blinking cursor at the clicked location.
            tag.IsShowingPlaceholderText = isShowingPlaceholderText;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertNode(tag);

            doc.Save(ArtifactsDir + "StructuredDocumentTag.PlaceholderBuildingBlock.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "StructuredDocumentTag.PlaceholderBuildingBlock.docx");
            tag = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            substituteBlock = (BuildingBlock)doc.GlossaryDocument.GetChild(NodeType.BuildingBlock, 0, true);

            Assert.AreEqual("Custom Placeholder", substituteBlock.Name);
            Assert.AreEqual(isShowingPlaceholderText, tag.IsShowingPlaceholderText);
            Assert.AreEqual(substituteBlock, tag.Placeholder);
            Assert.AreEqual(substituteBlock.Name, tag.PlaceholderName);
        }

        [Test]
        public void Lock()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.LockContentControl
            //ExFor:StructuredDocumentTag.LockContents
            //ExSummary:Shows how to restrict the editing of a StructuredDocumentTag.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a plain text StructuredDocumentTag of the PlainText type, which will function like a text box
            // It contains a default "Click here to enter text." prompt, which we can click and replace with our own text
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

            // We can prohibit the users from editing the inner text in Microsoft Word by setting this to true
            tag.LockContents = true;
            builder.Write("The contents of this StructuredDocumentTag cannot be edited: ");
            builder.InsertNode(tag);

            tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

            // Setting this to true will disable the deletion of this StructuredDocumentTag
            // by text editing operations in Microsoft Word
            tag.LockContentControl = true;

            builder.InsertParagraph();
            builder.Write("This StructuredDocumentTag cannot be deleted but its contents can be edited: ");
            builder.InsertNode(tag);

            doc.Save(ArtifactsDir + "StructuredDocumentTag.Lock.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "StructuredDocumentTag.Lock.docx");
            tag = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            Assert.True(tag.LockContents);
            Assert.False(tag.LockContentControl);

            tag = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 1, true);

            Assert.False(tag.LockContents);
            Assert.True(tag.LockContentControl);
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
            //ExFor:StructuredDocumentTag.ListItems
            //ExSummary:Shows how to work with StructuredDocumentTag nodes of the DropDownList type.
            // Create a blank document and insert a StructuredDocumentTag that will contain a drop-down list
            Document doc = new Document();
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Block);
            doc.FirstSection.Body.AppendChild(tag);

            // A drop-down list needs elements, each of which will be a SdtListItem
            SdtListItemCollection listItems = tag.ListItems;
            listItems.Add(new SdtListItem("Value 1"));

            // Each SdtListItem has text that will be displayed when the drop-down list is opened, and also a value
            // When we initialize with one string, we are providing just the value
            // Accordingly, value is passed as DisplayText and will consequently be displayed on the screen
            Assert.AreEqual(listItems[0].DisplayText, listItems[0].Value);

            // Add 3 more SdtListItems with non-empty strings passed to DisplayText
            listItems.Add(new SdtListItem("Item 2", "Value 2"));
            listItems.Add(new SdtListItem("Item 3", "Value 3"));
            listItems.Add(new SdtListItem("Item 4", "Value 4"));

            // We can obtain a count of the SdtListItems and also set the drop-down list's SelectedValue attribute to
            // automatically have one of them pre-selected when we open the document in Microsoft Word
            Assert.AreEqual(4, listItems.Count);
            listItems.SelectedValue = listItems[3];

            Assert.AreEqual("Value 4", listItems.SelectedValue.Value);

            // We can enumerate over the collection and print each element
            using (IEnumerator<SdtListItem> enumerator = listItems.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current != null)
                        Console.WriteLine($"List item: {enumerator.Current.DisplayText}, value: {enumerator.Current.Value}");
                }
            }

            // We can also remove elements one at a time
            listItems.RemoveAt(3);
            Assert.AreEqual(3, listItems.Count);

            // Make sure to update the SelectedValue's index if it ever ends up out of bounds before saving the document
            listItems.SelectedValue = listItems[1];
           
            doc.Save(ArtifactsDir + "StructuredDocumentTag.ListItemCollection.docx");

            // We can clear the whole collection at once too
            listItems.Clear();
            Assert.AreEqual(0, listItems.Count);
            //ExEnd
        }

        [Test]
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
            //ExFor:StructuredDocumentTag.XmlMapping
            //ExFor:XmlMapping.SetMapping(CustomXmlPart, String, String)
            //ExSummary:Shows how to create structured document tag with a custom XML data.
            Document doc = new Document();

            // Construct an XML part that contains data and add it to the document's collection
            // Once the "Developer" tab in Microsoft Word is enabled,
            // we can find elements from this collection as well as a couple defaults in the "XML Mapping Pane" 
            string xmlPartId = Guid.NewGuid().ToString("B");
            string xmlPartContent = "<root><text>Hello world!</text></root>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlPartContent);

            // The data we entered is stored in these attributes
            Assert.AreEqual(Encoding.ASCII.GetBytes(xmlPartContent), xmlPart.Data);
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
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            tag.XmlMapping.SetMapping(xmlPart, "/root[1]/text[1]", string.Empty);

            doc.FirstSection.Body.AppendChild(tag);

            doc.Save(ArtifactsDir + "StructuredDocumentTag.CustomXml.docx");
            //ExEnd

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "StructuredDocumentTag.CustomXml.docx", GoldsDir + "StructuredDocumentTag.CustomXml Gold.docx"));

            doc = new Document(ArtifactsDir + "StructuredDocumentTag.CustomXml.docx");
            xmlPart = doc.CustomXmlParts[0];

            Assert.True(Guid.TryParse(xmlPart.Id, out Guid temp));
            Assert.AreEqual("<root><text>Hello world!</text></root>", Encoding.UTF8.GetString(xmlPart.Data));
            Assert.AreEqual("http://www.w3.org/2001/XMLSchema", xmlPart.Schemas[0]);

            tag = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Assert.AreEqual("Hello world!", tag.GetText().Trim());
            Assert.AreEqual("/root[1]/text[1]", tag.XmlMapping.XPath);
            Assert.AreEqual(string.Empty, tag.XmlMapping.PrefixMappings);
        }

        [Test]
        public void XmlMapping()
        {
            //ExStart
            //ExFor:XmlMapping
            //ExFor:XmlMapping.CustomXmlPart
            //ExFor:XmlMapping.Delete
            //ExFor:XmlMapping.IsMapped
            //ExFor:XmlMapping.PrefixMappings
            //ExFor:XmlMapping.XPath
            //ExSummary:Shows how to set XML mappings for CustomXmlParts.
            Document doc = new Document();

            // Construct an XML part that contains data and add it to the document's CustomXmlPart collection
            string xmlPartId = Guid.NewGuid().ToString("B");
            string xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlPartContent);
            Console.WriteLine(Encoding.UTF8.GetString(xmlPart.Data));

            // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);

            // If we set a mapping for our StructuredDocumentTag,
            // it will only display a part of the CustomXmlPart that the XPath points to
            // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart
            tag.XmlMapping.SetMapping(xmlPart, "/root[1]/text[2]", "xmlns:ns='http://www.w3.org/2001/XMLSchema'");

            Assert.True(tag.XmlMapping.IsMapped);
            Assert.AreEqual(xmlPart, tag.XmlMapping.CustomXmlPart);
            Assert.AreEqual("/root[1]/text[2]", tag.XmlMapping.XPath);
            Assert.AreEqual("xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag.XmlMapping.PrefixMappings);

            // Add the StructuredDocumentTag to the document to display the content from our CustomXmlPart
            doc.FirstSection.Body.AppendChild(tag);
            doc.Save(ArtifactsDir + "StructuredDocumentTag.XmlMapping.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "StructuredDocumentTag.XmlMapping.docx");
            xmlPart = doc.CustomXmlParts[0];

            Assert.True(Guid.TryParse(xmlPart.Id, out Guid temp));
            Assert.AreEqual("<root><text>Text element #1</text><text>Text element #2</text></root>", Encoding.UTF8.GetString(xmlPart.Data));

            tag = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Assert.AreEqual("Text element #2", tag.GetText().Trim());
            Assert.AreEqual("/root[1]/text[2]", tag.XmlMapping.XPath);
            Assert.AreEqual("xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag.XmlMapping.PrefixMappings);
        }

        [Test]
        public void XmlMappingForStructuredDocumentTagRangeStart()
        {
            //ExStart
            //ExFor:StructuredDocumentTagRangeStart.XmlMapping
            //ExSummary:Shows how to set XML mappings for StructuredDocumentTagRangeStart.
            Document doc = new Document(MyDir + "Multi-section structured document tags.docx");

            // Construct an XML part that contains data and add it to the document's CustomXmlPart collection.
            string xmlPartId = Guid.NewGuid().ToString("B");
            string xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlPartContent);
            Console.WriteLine(Encoding.UTF8.GetString(xmlPart.Data));

            // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document.
            StructuredDocumentTagRangeStart sdtRangeStart = (StructuredDocumentTagRangeStart)doc.GetChild(NodeType.StructuredDocumentTagRangeStart, 0, true);

            // If we set a mapping for our StructuredDocumentTag,
            // it will only display a part of the CustomXmlPart that the XPath points to.
            // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
            sdtRangeStart.XmlMapping.SetMapping(xmlPart, "/root[1]/text[2]", null);

            doc.Save(ArtifactsDir + "StructuredDocumentTag.XmlMappingForStructuredDocumentTagRangeStart.docx");
            //ExEnd
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

            // Collections can be cloned, and elements can be added
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
            Document doc = new Document(MyDir + "Custom XML part in structured document tag.docx");

            // Structured document tags have IDs in the form of Guids
            StructuredDocumentTag tag = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Assert.AreEqual("{F3029283-4FF8-4DD2-9F31-395F19ACEE85}", tag.XmlMapping.StoreItemId);
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

            doc = DocumentHelper.SaveOpen(doc);

            StructuredDocumentTag sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
            Console.WriteLine("The Id of your custom xml part is: " + sdt.XmlMapping.StoreItemId);
        }

        [Test]
        public void ClearTextFromStructuredDocumentTags()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.Clear
            //ExSummary:Shows how to delete content of StructuredDocumentTag elements.
            Document doc = new Document();

            // Create a plain text structured document tag and append it to the document
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            doc.FirstSection.Body.AppendChild(tag);

            // This structured document tag, which is in the form of a text box, already displays placeholder text
            Assert.AreEqual("Click here to enter text.", tag.GetText().Trim());
            Assert.True(tag.IsShowingPlaceholderText);

            // Create a building block that 
            GlossaryDocument glossaryDoc = doc.GlossaryDocument;
            BuildingBlock substituteBlock = new BuildingBlock(glossaryDoc);
            substituteBlock.Name = "My placeholder";
            substituteBlock.AppendChild(new Section(glossaryDoc));
            substituteBlock.FirstSection.EnsureMinimum();
            substituteBlock.FirstSection.Body.FirstParagraph.AppendChild(new Run(glossaryDoc, "Custom placeholder text."));
            glossaryDoc.AppendChild(substituteBlock);

            // Set the tag's placeholder to the building block
            tag.PlaceholderName = "My placeholder";

            Assert.AreEqual("Custom placeholder text.", tag.GetText().Trim());
            Assert.True(tag.IsShowingPlaceholderText);

            // Edit the text of the structured document tag and disable showing of placeholder text
            Run run = (Run)tag.GetChild(NodeType.Run, 0, true);
            run.Text = "New text.";
            tag.IsShowingPlaceholderText = false;

            Assert.AreEqual("New text.", tag.GetText().Trim());

            tag.Clear();

            // Clearing a PlainText tag reverts these changes
            Assert.True(tag.IsShowingPlaceholderText);
            Assert.AreEqual("Custom placeholder text.", tag.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void AccessToBuildingBlockPropertiesFromDocPartObjSdt()
        {
            Document doc = new Document(MyDir + "Structured document tags with building blocks.docx");

            StructuredDocumentTag docPartObjSdt =
                (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            Assert.AreEqual(SdtType.DocPartObj, docPartObjSdt.SdtType);
            Assert.AreEqual("Table of Contents", docPartObjSdt.BuildingBlockGallery);
        }

        [Test]
        public void AccessToBuildingBlockPropertiesFromPlainTextSdt()
        {
            Document doc = new Document(MyDir + "Structured document tags with building blocks.docx");

            StructuredDocumentTag plainTextSdt =
                (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 1, true);

            Assert.AreEqual(SdtType.PlainText, plainTextSdt.SdtType);
            Assert.That(() => plainTextSdt.BuildingBlockGallery, Throws.TypeOf<InvalidOperationException>(),
                "BuildingBlockType is only accessible for BuildingBlockGallery SDT type.");
        }

        [Test]
        public void BuildingBlockCategories()
        {
            //ExStart
            //ExFor:StructuredDocumentTag.BuildingBlockCategory
            //ExFor:StructuredDocumentTag.BuildingBlockGallery
            //ExSummary:Shows how to insert a StructuredDocumentTag as a building block and set its category and gallery.
            Document doc = new Document();

            StructuredDocumentTag buildingBlockSdt =
                new StructuredDocumentTag(doc, SdtType.BuildingBlockGallery, MarkupLevel.Block)
                {
                    BuildingBlockCategory = "Built-in",
                    BuildingBlockGallery = "Table of Contents"
                };

            doc.FirstSection.Body.AppendChild(buildingBlockSdt);

            doc.Save(ArtifactsDir + "StructuredDocumentTag.BuildingBlockCategories.docx");
            //ExEnd

            buildingBlockSdt =
                (StructuredDocumentTag) doc.FirstSection.Body.GetChild(NodeType.StructuredDocumentTag, 0, true);

            Assert.AreEqual(SdtType.BuildingBlockGallery, buildingBlockSdt.SdtType);
            Assert.AreEqual("Table of Contents", buildingBlockSdt.BuildingBlockGallery);
            Assert.AreEqual("Built-in", buildingBlockSdt.BuildingBlockCategory);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void UpdateSdtContent(bool updateSdtContent)
        {
            //ExStart
            //ExFor:SaveOptions.UpdateSdtContent
            //ExSummary:Shows how structured document tags can be updated while saving to .pdf.
            Document doc = new Document();

            // Insert two StructuredDocumentTags; a date and a drop-down list 
            StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Block);
            tag.FullDate = DateTime.Now;

            doc.FirstSection.Body.AppendChild(tag);

            tag = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Block);
            tag.ListItems.Add(new SdtListItem("Value 1"));
            tag.ListItems.Add(new SdtListItem("Value 2"));
            tag.ListItems.Add(new SdtListItem("Value 3"));
            tag.ListItems.SelectedValue = tag.ListItems[1];

            doc.FirstSection.Body.AppendChild(tag);

            // We've selected default values for both tags
            // We can save those values in the document without immediately updating the tags, leaving them in their default state
            // by using a SaveOptions object with this flag set
            PdfSaveOptions options = new PdfSaveOptions();
            options.UpdateSdtContent = updateSdtContent;

            doc.Save(ArtifactsDir + "StructuredDocumentTag.UpdateSdtContent.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "StructuredDocumentTag.UpdateSdtContent.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.AreEqual(updateSdtContent ? "Value 2" : $"Click here to enter a date.{Environment.NewLine}Choose an item.",
                textAbsorber.Text);
#endif
        }

        [Test]
        public void FillTableUsingRepeatingSectionItem()
        {
            //ExStart
            //ExFor:SdtType
            //ExSummary:Shows how to fill the table with data contained in the XML part.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
 
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add("Books",
                "<books>" +
                "<book><title>Everyday Italian</title>" +
                "<author>Giada De Laurentiis</author></book>" +
                "<book><title>Harry Potter</title>" +
                "<author>J. K. Rowling</author></book>" +
                "<book><title>Learning XML</title>" +
                "<author>Erik T. Ray</author></book>" +
                "</books>");
 
            // Create headers for data from xml content
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Title");
            builder.InsertCell();
            builder.Write("Author");
            builder.EndRow();
            builder.EndTable();
 
            // Create table with RepeatingSection inside
            StructuredDocumentTag repeatingSectionSdt =
                new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Row);
            repeatingSectionSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book", string.Empty);
            table.AppendChild(repeatingSectionSdt);
 
            // Add RepeatingSectionItem inside RepeatingSection and mark it as a row
            StructuredDocumentTag repeatingSectionItemSdt =
                new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Row);
            repeatingSectionSdt.AppendChild(repeatingSectionItemSdt);
 
            Row row = new Row(doc);
            repeatingSectionItemSdt.AppendChild(row);
 
            // Map xml data with created table cells for book title and author
            StructuredDocumentTag titleSdt =
                new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
            titleSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/title[1]", string.Empty);
            row.AppendChild(titleSdt);
 
            StructuredDocumentTag authorSdt =
                new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Cell);
            authorSdt.XmlMapping.SetMapping(xmlPart, "/books[1]/book[1]/author[1]", string.Empty);
            row.AppendChild(authorSdt);
 
            doc.Save(ArtifactsDir + "StructuredDocumentTag.RepeatingSectionItem.docx");
			//ExEnd

            doc = new Document(ArtifactsDir + "StructuredDocumentTag.RepeatingSectionItem.docx");
            List<StructuredDocumentTag> tags = doc.GetChildNodes(NodeType.StructuredDocumentTag, true).OfType<StructuredDocumentTag>().ToList();

            Assert.AreEqual("/books[1]/book", tags[0].XmlMapping.XPath);
            Assert.AreEqual(string.Empty, tags[0].XmlMapping.PrefixMappings);

            Assert.AreEqual(string.Empty, tags[1].XmlMapping.XPath);
            Assert.AreEqual(string.Empty, tags[1].XmlMapping.PrefixMappings);

            Assert.AreEqual("/books[1]/book[1]/title[1]", tags[2].XmlMapping.XPath);
            Assert.AreEqual(string.Empty, tags[2].XmlMapping.PrefixMappings);

            Assert.AreEqual("/books[1]/book[1]/author[1]", tags[3].XmlMapping.XPath);
            Assert.AreEqual(string.Empty, tags[3].XmlMapping.PrefixMappings);

            Assert.AreEqual("Title\u0007Author\u0007\u0007" +
                            "Everyday Italian\u0007Giada De Laurentiis\u0007\u0007" +
                            "Harry Potter\u0007J. K. Rowling\u0007\u0007" +
                            "Learning XML\u0007Erik T. Ray\u0007\u0007", doc.FirstSection.Body.Tables[0].GetText().Trim());
        }

        [Test]
        public void CustomXmlPart()
        {
            // Obtain an XML in the form of a string
            string xmlString = "<?xml version=\"1.0\"?>" +
                               "<Company>" +
                               "<Employee id=\"1\">" +
                               "<FirstName>John</FirstName>" +
                               "<LastName>Doe</LastName>" +
                               "</Employee>" +
                               "<Employee id=\"2\">" +
                               "<FirstName>Jane</FirstName>" +
                               "<LastName>Doe</LastName>" +
                               "</Employee>" +
                               "</Company>";

            Document doc = new Document();

            // Insert the full XML document as a custom document part
            // The mapping for this part will be seen in the "XML Mapping Pane" in the "Developer" tab, if it is enabled
            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(Guid.NewGuid().ToString("B"), xmlString);

            // None of the XML is in the document body at this point
            // Create a StructuredDocumentTag, which will refer to a single element from the XML with an XPath
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
            sdt.XmlMapping.SetMapping(xmlPart, "Company//Employee[@id='2']/FirstName", "");

            // Add the StructuredDocumentTag to the document to display the element in the text 
            doc.FirstSection.Body.AppendChild(sdt);
        }

        [Test]
        public void MultiSectionTags()
        {
            //ExStart
            //ExFor:StructuredDocumentTagRangeStart
            //ExFor:StructuredDocumentTagRangeStart.Id
            //ExFor:StructuredDocumentTagRangeStart.Title
            //ExFor:StructuredDocumentTagRangeStart.IsShowingPlaceholderText
            //ExFor:StructuredDocumentTagRangeStart.LockContentControl
            //ExFor:StructuredDocumentTagRangeStart.LockContents
            //ExFor:StructuredDocumentTagRangeStart.Level
            //ExFor:StructuredDocumentTagRangeStart.RangeEnd
            //ExFor:StructuredDocumentTagRangeStart.SdtType
            //ExFor:StructuredDocumentTagRangeStart.Tag
            //ExFor:StructuredDocumentTagRangeEnd
            //ExFor:StructuredDocumentTagRangeEnd.Id
            //ExSummary:Shows how to get multi-section structured document tags properties.
            Document doc = new Document(MyDir + "Multi-section structured document tags.docx");

            // Note that these nodes can be a child of NodeType.Body node only and all properties of these nodes are read-only.
            StructuredDocumentTagRangeStart rangeStartTag =
                doc.GetChildNodes(NodeType.StructuredDocumentTagRangeStart, true)[0] as StructuredDocumentTagRangeStart;
            StructuredDocumentTagRangeEnd rangeEndTag =
                doc.GetChildNodes(NodeType.StructuredDocumentTagRangeEnd, true)[0] as StructuredDocumentTagRangeEnd;

            Assert.AreEqual(rangeStartTag.Id, rangeEndTag.Id); //ExSkip
            Assert.AreEqual(NodeType.StructuredDocumentTagRangeStart, rangeStartTag.NodeType); //ExSkip
            Assert.AreEqual(NodeType.StructuredDocumentTagRangeEnd, rangeEndTag.NodeType); //ExSkip

            Console.WriteLine("StructuredDocumentTagRangeStart values:");
            Console.WriteLine($"\t|Id: {rangeStartTag.Id}");
            Console.WriteLine($"\t|Title: {rangeStartTag.Title}");
            Console.WriteLine($"\t|IsShowingPlaceholderText: {rangeStartTag.IsShowingPlaceholderText}");
            Console.WriteLine($"\t|LockContentControl: {rangeStartTag.LockContentControl}");
            Console.WriteLine($"\t|LockContents: {rangeStartTag.LockContents}");
            Console.WriteLine($"\t|Level: {rangeStartTag.Level}");
            Console.WriteLine($"\t|NodeType: {rangeStartTag.NodeType}");
            Console.WriteLine($"\t|RangeEnd: {rangeStartTag.RangeEnd}");
            Console.WriteLine($"\t|SdtType: {rangeStartTag.SdtType}");
            Console.WriteLine($"\t|Tag: {rangeStartTag.Tag}\n");

            Console.WriteLine("StructuredDocumentTagRangeEnd values:");
            Console.WriteLine($"\t|Id: {rangeEndTag.Id}");
            Console.WriteLine($"\t|NodeType: {rangeEndTag.NodeType}");
            //ExEnd
        }
    }
}