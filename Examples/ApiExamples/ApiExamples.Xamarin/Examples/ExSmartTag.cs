using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExSmartTag : ApiExampleBase
    {
        //ExStart
        //ExFor:CompositeNode.RemoveSmartTags
        //ExFor:CustomXmlProperty
        //ExFor:CustomXmlProperty.#ctor(String,String,String)
        //ExFor:CustomXmlProperty.Name
        //ExFor:CustomXmlProperty.Value
        //ExFor:Markup.SmartTag
        //ExFor:Markup.SmartTag.#ctor(DocumentBase)
        //ExFor:Markup.SmartTag.Accept(DocumentVisitor)
        //ExFor:Markup.SmartTag.Element
        //ExFor:Markup.SmartTag.Properties
        //ExFor:Markup.SmartTag.Uri
        //ExSummary:Shows how to create smart tags.
        [Test] //ExSkip
        public void Create()
        {
            Document doc = new Document();

            // A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
            // such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
            SmartTag smartTag = new SmartTag(doc);

            // Smart tags are composite nodes that contain their recognized text in its entirety.
            // Add contents to this smart tag manually.
            smartTag.AppendChild(new Run(doc, "May 29, 2019"));

            // Microsoft Word may recognize the above contents as being a date.
            // Smart tags use the "Element" property to reflect the type of data they contain.
            smartTag.Element = "date";

            // Some smart tag types process their contents further into custom XML properties.
            smartTag.Properties.Add(new CustomXmlProperty("Day", string.Empty, "29"));
            smartTag.Properties.Add(new CustomXmlProperty("Month", string.Empty, "5"));
            smartTag.Properties.Add(new CustomXmlProperty("Year", string.Empty, "2019"));

            // Set the smart tag's URI to the default value.
            smartTag.Uri = "urn:schemas-microsoft-com:office:smarttags";

            doc.FirstSection.Body.FirstParagraph.AppendChild(smartTag);
            doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, " is a date. "));

            // Create another smart tag for a stock ticker.
            smartTag = new SmartTag(doc);
            smartTag.Element = "stockticker";
            smartTag.Uri = "urn:schemas-microsoft-com:office:smarttags";

            smartTag.AppendChild(new Run(doc, "MSFT"));

            doc.FirstSection.Body.FirstParagraph.AppendChild(smartTag);
            doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, " is a stock ticker."));

            // Print all the smart tags in our document using a document visitor.
            doc.Accept(new SmartTagPrinter());

            // Older versions of Microsoft Word support smart tags.
            doc.Save(ArtifactsDir + "SmartTag.Create.doc");

            // Use the "RemoveSmartTags" method to remove all smart tags from a document.
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.SmartTag, true).Count);

            doc.RemoveSmartTags();

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.SmartTag, true).Count);
            TestCreate(new Document(ArtifactsDir + "SmartTag.Create.doc")); //ExSkip
        }

        /// <summary>
        /// Prints visited smart tags and their contents.
        /// </summary>
        private class SmartTagPrinter : DocumentVisitor
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
                        properties[index++] = $"\"{cxp.Name}\" = \"{cxp.Value}\"";

                    Console.WriteLine(string.Join(", ", properties));
                }

                return VisitorAction.Continue;
            }
        }
        //ExEnd

        public void TestCreate(Document doc)
        {
            SmartTag smartTag = (SmartTag)doc.GetChild(NodeType.SmartTag, 0, true);

            Assert.AreEqual("date", smartTag.Element);
            Assert.AreEqual("May 29, 2019", smartTag.GetText());
            Assert.AreEqual("urn:schemas-microsoft-com:office:smarttags", smartTag.Uri);

            Assert.AreEqual("Day", smartTag.Properties[0].Name);
            Assert.AreEqual(string.Empty, smartTag.Properties[0].Uri);
            Assert.AreEqual("29", smartTag.Properties[0].Value);
            Assert.AreEqual("Month", smartTag.Properties[1].Name);
            Assert.AreEqual(string.Empty, smartTag.Properties[1].Uri);
            Assert.AreEqual("5", smartTag.Properties[1].Value);
            Assert.AreEqual("Year", smartTag.Properties[2].Name);
            Assert.AreEqual(string.Empty, smartTag.Properties[2].Uri);
            Assert.AreEqual("2019", smartTag.Properties[2].Value);

            smartTag = (SmartTag)doc.GetChild(NodeType.SmartTag, 1, true);

            Assert.AreEqual("stockticker", smartTag.Element);
            Assert.AreEqual("MSFT", smartTag.GetText());
            Assert.AreEqual("urn:schemas-microsoft-com:office:smarttags", smartTag.Uri);
            Assert.AreEqual(0, smartTag.Properties.Count);
        }

        [Test]
        public void Properties()
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
            Document doc = new Document(MyDir + "Smart tags.doc");

            // A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
            // such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
            // In Word 2003, we can enable smart tags via "Tools" -> "AutoCorrect options..." -> "SmartTags".
            // In our input document, there are three objects that Microsoft Word registered as smart tags.
            // Smart tags may be nested, so this collection contains more.
            SmartTag[] smartTags = doc.GetChildNodes(NodeType.SmartTag, true).OfType<SmartTag>().ToArray();

            Assert.AreEqual(8, smartTags.Length);

            // The "Properties" member of a smart tag contains its metadata, which will be different for each type of smart tag.
            // The properties of a "date"-type smart tag contain its year, month, and day.
            CustomXmlPropertyCollection properties = smartTags[7].Properties;

            Assert.AreEqual(4, properties.Count);

            using (IEnumerator<CustomXmlProperty> enumerator = properties.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"Property name: {enumerator.Current.Name}, value: {enumerator.Current.Value}");
                    Assert.AreEqual("", enumerator.Current.Uri);
                }
            }

            // We can also access the properties in various ways, such as a key-value pair.
            Assert.True(properties.Contains("Day"));
            Assert.AreEqual("22", properties["Day"].Value);
            Assert.AreEqual("2003", properties[2].Value);
            Assert.AreEqual(1, properties.IndexOfKey("Month"));

            // Below are three ways of removing elements from the properties collection.
            // 1 -  Remove by index:
            properties.RemoveAt(3);

            Assert.AreEqual(3, properties.Count);

            // 2 -  Remove by name:
            properties.Remove("Year");

            Assert.AreEqual(2, properties.Count);

            // 3 -  Clear the entire collection at once:
            properties.Clear();

            Assert.AreEqual(0, properties.Count);
            //ExEnd
        }
    }
}
