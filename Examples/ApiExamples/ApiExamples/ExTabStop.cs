// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Linq;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTabStop : ApiExampleBase
    {
        [Test]
        public void AddTabStops()
        {
            //ExStart
            //ExFor:TabStopCollection.Add(TabStop)
            //ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
            //ExSummary:Shows how to add custom tab stops to a document.
            Document doc = new Document();
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            // Below are two ways of adding tab stops to a paragraph's collection of tab stops via the "ParagraphFormat" property.
            // 1 -  Create a "TabStop" object, and then add it to the collection:
            TabStop tabStop = new TabStop(ConvertUtil.InchToPoint(3), TabAlignment.Left, TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(tabStop);

            // 2 -  Pass the values for properties of a new tab stop to the "Add" method:
            paragraph.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(100), TabAlignment.Left,
                TabLeader.Dashes);

            // Add tab stops at 5 cm to all paragraphs.
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
            {
                para.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(50), TabAlignment.Left,
                    TabLeader.Dashes);
            }

            // Every "tab" character takes the builder's cursor to the location of the next tab stop.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Start\tTab 1\tTab 2\tTab 3\tTab 4");

            doc.Save(ArtifactsDir + "TabStopCollection.AddTabStops.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "TabStopCollection.AddTabStops.docx");
            TabStopCollection tabStops = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.TabStops;

            TestUtil.VerifyTabStop(141.75d, TabAlignment.Left, TabLeader.Dashes, false, tabStops[0]);
            TestUtil.VerifyTabStop(216.0d, TabAlignment.Left, TabLeader.Dashes, false, tabStops[1]);
            TestUtil.VerifyTabStop(283.45d, TabAlignment.Left, TabLeader.Dashes, false, tabStops[2]);
        }

        [Test]
        public void TabStopCollection()
        {
            //ExStart
            //ExFor:TabStop.#ctor
            //ExFor:TabStop.#ctor(Double)
            //ExFor:TabStop.#ctor(Double,TabAlignment,TabLeader)
            //ExFor:TabStop.Equals(TabStop)
            //ExFor:TabStop.IsClear
            //ExFor:TabStopCollection
            //ExFor:TabStopCollection.After(Double)
            //ExFor:TabStopCollection.Before(Double)
            //ExFor:TabStopCollection.Clear
            //ExFor:TabStopCollection.Count
            //ExFor:TabStopCollection.Equals(TabStopCollection)
            //ExFor:TabStopCollection.Equals(Object)
            //ExFor:TabStopCollection.GetHashCode
            //ExFor:TabStopCollection.Item(Double)
            //ExFor:TabStopCollection.Item(Int32)
            //ExSummary:Shows how to work with a document's collection of tab stops.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            TabStopCollection tabStops = builder.ParagraphFormat.TabStops;

            // 72 points is one "inch" on the Microsoft Word tab stop ruler.
            tabStops.Add(new TabStop(72.0));
            tabStops.Add(new TabStop(432.0, TabAlignment.Right, TabLeader.Dashes));

            Assert.AreEqual(2, tabStops.Count);
            Assert.IsFalse(tabStops[0].IsClear);
            Assert.IsFalse(tabStops[0].Equals(tabStops[1]));

            // Every "tab" character takes the builder's cursor to the location of the next tab stop.
            builder.Writeln("Start\tTab 1\tTab 2");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(2, paragraphs.Count);

            // Each paragraph gets its tab stop collection, which clones its values from the document builder's tab stop collection.
            Assert.AreEqual(paragraphs[0].ParagraphFormat.TabStops, paragraphs[1].ParagraphFormat.TabStops);
            Assert.AreNotSame(paragraphs[0].ParagraphFormat.TabStops, paragraphs[1].ParagraphFormat.TabStops);

            // A tab stop collection can point us to TabStops before and after certain positions.
            Assert.AreEqual(72.0, tabStops.Before(100.0).Position);
            Assert.AreEqual(432.0, tabStops.After(100.0).Position);

            // We can clear a paragraph's tab stop collection to revert to the default tabbing behavior.
            paragraphs[1].ParagraphFormat.TabStops.Clear();

            Assert.AreEqual(0, paragraphs[1].ParagraphFormat.TabStops.Count);

            doc.Save(ArtifactsDir + "TabStopCollection.TabStopCollection.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "TabStopCollection.TabStopCollection.docx");
            tabStops = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.TabStops;

            Assert.AreEqual(2, tabStops.Count);
            TestUtil.VerifyTabStop(72.0d, TabAlignment.Left, TabLeader.None, false, tabStops[0]);
            TestUtil.VerifyTabStop(432.0d, TabAlignment.Right, TabLeader.Dashes, false, tabStops[1]);

            tabStops = doc.FirstSection.Body.Paragraphs[1].ParagraphFormat.TabStops;

            Assert.AreEqual(0, tabStops.Count);
        }

        [Test]
        public void RemoveByIndex()
        {
            //ExStart
            //ExFor:TabStopCollection.RemoveByIndex
            //ExSummary:Shows how to select a tab stop in a document by its index and remove it.
            Document doc = new Document();
            TabStopCollection tabStops = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.TabStops;

            tabStops.Add(ConvertUtil.MillimeterToPoint(30), TabAlignment.Left, TabLeader.Dashes);
            tabStops.Add(ConvertUtil.MillimeterToPoint(60), TabAlignment.Left, TabLeader.Dashes);

            Assert.AreEqual(2, tabStops.Count);

            // Remove the first tab stop.
            tabStops.RemoveByIndex(0);

            Assert.AreEqual(1, tabStops.Count);

            doc.Save(ArtifactsDir + "TabStopCollection.RemoveByIndex.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "TabStopCollection.RemoveByIndex.docx");

            TestUtil.VerifyTabStop(170.1d, TabAlignment.Left, TabLeader.Dashes, false, doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.TabStops[0]);
        }

        [Test]
        public void GetPositionByIndex()
        {
            //ExStart
            //ExFor:TabStopCollection.GetPositionByIndex
            //ExSummary:Shows how to find a tab, stop by its index and verify its position.
            Document doc = new Document();
            TabStopCollection tabStops = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.TabStops;

            tabStops.Add(ConvertUtil.MillimeterToPoint(30), TabAlignment.Left, TabLeader.Dashes);
            tabStops.Add(ConvertUtil.MillimeterToPoint(60), TabAlignment.Left, TabLeader.Dashes);

            // Verify the position of the second tab stop in the collection.
            Assert.AreEqual(ConvertUtil.MillimeterToPoint(60), tabStops.GetPositionByIndex(1), 0.1d);
            //ExEnd
        }

        [Test]
        public void GetIndexByPosition()
        {
            //ExStart
            //ExFor:TabStopCollection.GetIndexByPosition
            //ExSummary:Shows how to look up a position to see if a tab stop exists there and obtain its index.
            Document doc = new Document();
            TabStopCollection tabStops = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.TabStops;

            // Add a tab stop at a position of 30mm.
            tabStops.Add(ConvertUtil.MillimeterToPoint(30), TabAlignment.Left, TabLeader.Dashes);

            // A result of "0" returned by "GetIndexByPosition" confirms that a tab stop
            // at 30mm exists in this collection, and it is at index 0.
            Assert.AreEqual(0, tabStops.GetIndexByPosition(ConvertUtil.MillimeterToPoint(30)));

            // A "-1" returned by "GetIndexByPosition" confirms that
            // there is no tab stop in this collection with a position of 60mm.
            Assert.AreEqual(-1, tabStops.GetIndexByPosition(ConvertUtil.MillimeterToPoint(60)));
            //ExEnd
        }
    }
}