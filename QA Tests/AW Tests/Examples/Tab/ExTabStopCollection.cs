// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Tab
{
    [TestFixture]
    public class ExTabStopCollection : QaTestsBase
    {
        [Test]
        public void ClearEx()
        {
            //ExStart
            //ExFor:TabStopCollection.Clear
            //ExSummary:Shows how to remove all tab stops from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.TableOfContents.doc");

            // Clear all tab stops from every paragraph.
            foreach (Aspose.Words.Paragraph para in doc.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
            {
                para.ParagraphFormat.TabStops.Clear();
            }

            doc.Save(ExDir + "Document.AllTabStopsRemoved Out.doc");
            //ExEnd
        }

        [Test]
        public void AddEx()
        {
            //ExStart
            //ExFor:TabStopCollection.Add(TabStop)
            //ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
            //ExSummary:Shows how to create tab stops and add them to a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.Paragraph paragraph = (Aspose.Words.Paragraph)doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, true);

            // Create a TabStop object and add it to the document.
            Aspose.Words.TabStop tabStop = new Aspose.Words.TabStop(Aspose.Words.ConvertUtil.InchToPoint(3), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(tabStop);

            // Add a tab stop without explicitly creating new TabStop objects.
            paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(100), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);

            // Add tab stops at 5 cm to all paragraphs.
            foreach (Aspose.Words.Paragraph para in doc.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
            {
                para.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(50), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);
            }

            doc.Save(ExDir + "Document.AddedTabStops Out.doc");
            //ExEnd
        }

        [Test]
        public void RemoveByIndexEx()
        {
            //ExStart
            //ExFor:TabStopCollection.RemoveByIndex
            //ExSummary:Shows how to select a tab stop in a document by it's index and remove it.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.Paragraph paragraph = (Aspose.Words.Paragraph)doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, true);

            paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(30), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(60), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);

            // Tab stop placed at 30 mm is removed
            paragraph.ParagraphFormat.TabStops.RemoveByIndex(0);

            Console.WriteLine(paragraph.ParagraphFormat.TabStops.Count);

            doc.Save(ExDir + "Document.RemovedTabStopsByIndex Out.doc");
            //ExEnd
        }

        [Test]
        public void GetPositionByIndexEx()
        {
            //ExStart
            //ExFor:TabStopCollection.GetPositionByIndex
            //ExSummary:Shows how to find a tab stop by it's index and get its position.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.Paragraph paragraph = (Aspose.Words.Paragraph)doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, true);

            paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(30), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(60), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);

            Console.WriteLine("Tab stop at index {0} of the first paragraph is at {1} points.", 1, paragraph.ParagraphFormat.TabStops.GetPositionByIndex(1));
            //ExEnd
        }

        [Test]
        public void GetIndexByPositionEx()
        {
            //ExStart
            //ExFor:TabStopCollection.GetIndexByPosition
            //ExSummary:Shows how to look up a position to see if a tab stop exists there, and if so, obtain its index.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.Paragraph paragraph = (Aspose.Words.Paragraph)doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, true);

            paragraph.ParagraphFormat.TabStops.Add(Aspose.Words.ConvertUtil.MillimeterToPoint(30), Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);

            // An output of -1 signifies that there is no tab stop at that position.
            Console.WriteLine(paragraph.ParagraphFormat.TabStops.GetIndexByPosition(Aspose.Words.ConvertUtil.MillimeterToPoint(30))); // 0
            Console.WriteLine(paragraph.ParagraphFormat.TabStops.GetIndexByPosition(Aspose.Words.ConvertUtil.MillimeterToPoint(60))); // -1
            //ExEnd
        }
    }
}
