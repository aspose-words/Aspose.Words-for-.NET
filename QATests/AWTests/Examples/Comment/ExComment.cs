// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Comment
{
    [TestFixture]
    public class ExComment : QaTestsBase
    {
        [Test]
        public void SetTextEx()
        {
            //ExStart
            //ExFor:Comment.SetText
            //ExSummary:Shows how to add a comment to a document and set it's text.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.Comment comment = new Aspose.Words.Comment(doc, "John Doe", "J.D.", DateTime.Now);
            builder.CurrentParagraph.AppendChild(comment);
            comment.SetText("My comment.");
            //ExEnd
        }
    }
}
