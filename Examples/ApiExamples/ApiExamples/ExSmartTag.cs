// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
        //ExFor:SmartTag
        //ExFor:SmartTag.#ctor(DocumentBase)
        //ExFor:SmartTag.Accept(DocumentVisitor)
        //ExFor:SmartTag.AcceptStart(DocumentVisitor)
        //ExFor:SmartTag.AcceptEnd(DocumentVisitor)
        //ExFor:SmartTag.Element
        //ExFor:SmartTag.Properties
        //ExFor:SmartTag.Uri
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
            Assert.That(doc.GetChildNodes(NodeType.SmartTag, true).Count, Is.EqualTo(2));

            doc.RemoveSmartTags();

            Assert.That(doc.GetChildNodes(NodeType.SmartTag, true).Count, Is.EqualTo(0));
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

            Assert.That(smartTag.Element, Is.EqualTo("date"));
            Assert.That(smartTag.GetText(), Is.EqualTo("May 29, 2019"));
            Assert.That(smartTag.Uri, Is.EqualTo("urn:schemas-microsoft-com:office:smarttags"));

            Assert.That(smartTag.Properties[0].Name, Is.EqualTo("Day"));
            Assert.That(smartTag.Properties[0].Uri, Is.EqualTo(string.Empty));
            Assert.That(smartTag.Properties[0].Value, Is.EqualTo("29"));
            Assert.That(smartTag.Properties[1].Name, Is.EqualTo("Month"));
            Assert.That(smartTag.Properties[1].Uri, Is.EqualTo(string.Empty));
            Assert.That(smartTag.Properties[1].Value, Is.EqualTo("5"));
            Assert.That(smartTag.Properties[2].Name, Is.EqualTo("Year"));
            Assert.That(smartTag.Properties[2].Uri, Is.EqualTo(string.Empty));
            Assert.That(smartTag.Properties[2].Value, Is.EqualTo("2019"));

            smartTag = (SmartTag)doc.GetChild(NodeType.SmartTag, 1, true);

            Assert.That(smartTag.Element, Is.EqualTo("stockticker"));
            Assert.That(smartTag.GetText(), Is.EqualTo("MSFT"));
            Assert.That(smartTag.Uri, Is.EqualTo("urn:schemas-microsoft-com:office:smarttags"));
            Assert.That(smartTag.Properties.Count, Is.EqualTo(0));
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

            Assert.That(smartTags.Length, Is.EqualTo(8));

            // The "Properties" member of a smart tag contains its metadata, which will be different for each type of smart tag.
            // The properties of a "date"-type smart tag contain its year, month, and day.
            CustomXmlPropertyCollection properties = smartTags[7].Properties;

            Assert.That(properties.Count, Is.EqualTo(4));

            using (IEnumerator<CustomXmlProperty> enumerator = properties.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"Property name: {enumerator.Current.Name}, value: {enumerator.Current.Value}");
                    Assert.That(enumerator.Current.Uri, Is.EqualTo(""));
                }
            }

            // We can also access the properties in various ways, such as a key-value pair.
            Assert.That(properties.Contains("Day"), Is.True);
            Assert.That(properties["Day"].Value, Is.EqualTo("22"));
            Assert.That(properties[2].Value, Is.EqualTo("2003"));
            Assert.That(properties.IndexOfKey("Month"), Is.EqualTo(1));

            // Below are three ways of removing elements from the properties collection.
            // 1 -  Remove by index:
            properties.RemoveAt(3);

            Assert.That(properties.Count, Is.EqualTo(3));

            // 2 -  Remove by name:
            properties.Remove("Year");

            Assert.That(properties.Count, Is.EqualTo(2));

            // 3 -  Clear the entire collection at once:
            properties.Clear();

            Assert.That(properties.Count, Is.EqualTo(0));
            //ExEnd
        }
    }
}
