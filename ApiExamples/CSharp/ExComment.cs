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
    public class ExComment : ApiExampleBase
    {
        [Test]
        public void SetTextEx()
        {
            //ExStart
            //ExFor:Comment.SetText
            //ExSummary:Shows how to add a comment to a document and set it's text.
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            builder.CurrentParagraph.AppendChild(comment);
            comment.SetText("My comment.");
            //ExEnd
        }
    }
}
