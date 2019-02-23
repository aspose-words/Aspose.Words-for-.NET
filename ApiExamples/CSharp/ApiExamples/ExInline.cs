// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
    class ExInline : ApiExampleBase
    {
        [Test]
        public void Inline()
        {
            //ExStart
            //ExFor:Inline
            //ExFor:Inline.IsDeleteRevision
            //ExFor:Inline.IsFormatRevision
            //ExFor:Inline.IsInsertRevision
            //ExFor:Inline.IsMoveFromRevision
            //ExFor:Inline.IsMoveToRevision
            //ExFor:Inline.ParentParagraph
            //ExSummary:.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Is not a revision");

            Assert.False(doc.HasRevisions);
            Assert.False(builder.CurrentParagraph.Runs[0].IsInsertRevision);

            doc.StartTrackRevisions("John Doe", DateTime.Now);

            builder.Write("Is a revision");
            Assert.True(doc.HasRevisions);
            Assert.True(builder.CurrentParagraph.Runs[1].IsInsertRevision);
        }
    }
}
