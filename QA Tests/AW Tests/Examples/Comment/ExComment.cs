﻿// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Comment
{
    [TestFixture]
    public class ExComment : QaTestsBase
    {
        [Test]
        public void AcceptAllRevisions()
        {
            //ExStart
            //ExFor:Document.AcceptAllRevisions
            //ExId:AcceptAllRevisions
            //ExSummary:Shows how to accept all tracking changes in the document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            doc.AcceptAllRevisions();
            //ExEnd
        }
    }
}
