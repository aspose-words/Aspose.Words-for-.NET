// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;

using Aspose.Words;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTabStopCollection : ApiExampleBase
    {
        [Test]
        public void ClearEx()
        {
            //ExStart
            //ExFor:TabStopCollection.Clear
            //ExSummary:Shows how to remove all tab stops from a document.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            // Clear all tab stops from every paragraph.
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.TabStops.Clear();
            }

            doc.Save(MyDir + @"\Artifacts\Document.AllTabStopsRemoved.doc");
            //ExEnd
        }

        [Test]
        public void AddEx()
        {
            //ExStart
            //ExFor:TabStopCollection.Add(TabStop)
            //ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
            //ExSummary:Shows how to create tab stops and add them to a document.
            Document doc = new Document(MyDir + "Document.doc");
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            // Create a TabStop object and add it to the document.
            TabStop tabStop = new TabStop(ConvertUtil.InchToPoint(3), TabAlignment.Left, TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(tabStop);

            // Add a tab stop without explicitly creating new TabStop objects.
            paragraph.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(100), TabAlignment.Left, TabLeader.Dashes);

            // Add tab stops at 5 cm to all paragraphs.
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(50), TabAlignment.Left, TabLeader.Dashes);
            }

            doc.Save(MyDir + @"\Artifacts\Document.AddedTabStops.doc");
            //ExEnd
        }

        [Test]
        public void RemoveByIndexEx()
        {
            //ExStart
            //ExFor:TabStopCollection.RemoveByIndex
            //ExSummary:Shows how to select a tab stop in a document by its index and remove it.
            Document doc = new Document(MyDir + "Document.doc");
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            paragraph.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(30), TabAlignment.Left, TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(60), TabAlignment.Left, TabLeader.Dashes);

            // Tab stop placed at 30 mm is removed
            paragraph.ParagraphFormat.TabStops.RemoveByIndex(0);

            Console.WriteLine(paragraph.ParagraphFormat.TabStops.Count);

            doc.Save(MyDir + @"\Artifacts\Document.RemovedTabStopsByIndex.doc");
            //ExEnd
        }

        [Test]
        public void GetPositionByIndexEx()
        {
            //ExStart
            //ExFor:TabStopCollection.GetPositionByIndex
            //ExSummary:Shows how to find a tab stop by it's index and get its position.
            Document doc = new Document(MyDir + "Document.doc");
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            paragraph.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(30), TabAlignment.Left, TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(60), TabAlignment.Left, TabLeader.Dashes);

            Console.WriteLine("Tab stop at index {0} of the first paragraph is at {1} points.", 1, paragraph.ParagraphFormat.TabStops.GetPositionByIndex(1));
            //ExEnd
        }

        [Test]
        public void GetIndexByPositionEx()
        {
            //ExStart
            //ExFor:TabStopCollection.GetIndexByPosition
            //ExSummary:Shows how to look up a position to see if a tab stop exists there, and if so, obtain its index.
            Document doc = new Document(MyDir + "Document.doc");
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            paragraph.ParagraphFormat.TabStops.Add(ConvertUtil.MillimeterToPoint(30), TabAlignment.Left, TabLeader.Dashes);

            // An output of -1 signifies that there is no tab stop at that position.
            Console.WriteLine(paragraph.ParagraphFormat.TabStops.GetIndexByPosition(ConvertUtil.MillimeterToPoint(30))); // 0
            Console.WriteLine(paragraph.ParagraphFormat.TabStops.GetIndexByPosition(ConvertUtil.MillimeterToPoint(60))); // -1
            //ExEnd
        }
    }
}
