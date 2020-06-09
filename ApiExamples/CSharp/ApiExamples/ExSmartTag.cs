using System;
using System.Collections.Generic;
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
        //ExFor:Markup.SmartTag.#ctor(Aspose.Words.DocumentBase)
        //ExFor:Markup.SmartTag.Accept(Aspose.Words.DocumentVisitor)
        //ExFor:Markup.SmartTag.Element
        //ExFor:Markup.SmartTag.Properties
        //ExFor:Markup.SmartTag.Uri
        //ExSummary:Shows how to create smart tags.
        [Test] //ExSkip
        public void Create()
        {
            Document doc = new Document();
            SmartTag smartTag = new SmartTag(doc);
            smartTag.Element = "date";

            // Specify a date and set smart tag properties accordingly
            smartTag.AppendChild(new Run(doc, "May 29, 2019"));

            smartTag.Properties.Add(new CustomXmlProperty("Day", string.Empty, "29"));
            smartTag.Properties.Add(new CustomXmlProperty("Month", string.Empty, "5"));
            smartTag.Properties.Add(new CustomXmlProperty("Year", string.Empty, "2019"));

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

            // SmartTags are supported by older versions of microsoft Word
            doc.Save(ArtifactsDir + "SmartTag.Create.doc");

            // We can strip a document of all its smart tags with RemoveSmartTags()
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.SmartTag, true).Count);
            doc.RemoveSmartTags();
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.SmartTag, true).Count);
            TestCreate(new Document(ArtifactsDir + "SmartTag.Create.doc")); //ExSkip
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
            // Open a document that contains smart tags and their collection
            Document doc = new Document(MyDir + "Smart tags.doc");

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
    }
}
