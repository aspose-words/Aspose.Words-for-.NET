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
    }
}
