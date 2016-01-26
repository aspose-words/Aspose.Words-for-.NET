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
        public void ClearAllAttrsEx()
        {
            //ExStart
            //ExFor:TabStopCollection.Clear
            //ExSummary:Shows how to clear a document of all tab stops.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.TableOfContents.doc");

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
            //ExSummary:Shows how to create and add tabStop objects to a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.Paragraph paragraph = (Aspose.Words.Paragraph)doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, true);

            // Create a TabStop object and add it into the document
            Aspose.Words.TabStop tabStop = new Aspose.Words.TabStop(84.99, Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(tabStop);

            // Add a TabStop without explicitly initializing  
            paragraph.ParagraphFormat.TabStops.Add(169.98, Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);

            doc.Save(ExDir + "Document.AddedTabStops Out.doc");
            //ExEnd
        }

        [Test]
        public void RemoveByIndexEx()
        {
            //ExStart
            //ExFor:TabStopCollection.RemoveByIndex
            //ExSummary:Shows how to select a TabStop in a document by it's index and remove it.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.Paragraph paragraph = (Aspose.Words.Paragraph)doc.GetChild(Aspose.Words.NodeType.Paragraph, 0, true);

            paragraph.ParagraphFormat.TabStops.Add(84.99, Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);
            paragraph.ParagraphFormat.TabStops.Add(169.98, Aspose.Words.TabAlignment.Left, Aspose.Words.TabLeader.Dashes);

            // TabStop at 84.99 points is removed
            paragraph.ParagraphFormat.TabStops.RemoveByIndex(0);

            doc.Save(ExDir + "Document.RemovedTabStopsByIndex Out.doc");
            //ExEnd
        }
    }
}
