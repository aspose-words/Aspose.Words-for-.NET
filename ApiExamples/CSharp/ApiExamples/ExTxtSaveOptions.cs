// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTxtSaveOptions : ApiExampleBase
    {
        [Test]
        public void PageBreaks()
        {
            //ExStart
            //ExFor:TxtSaveOptions.ForcePageBreaks
            //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
            Document doc = new Document(MyDir + "SaveOptions.PageBreaks.docx");

            TxtSaveOptions saveOptions = new TxtSaveOptions { ForcePageBreaks = false };

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PageBreaks.txt", saveOptions);
            //ExEnd
        }

        [Test]
        public void AddBidiMarks()
        {
            //ExStart
            //ExFor:TxtSaveOptions.AddBidiMarks
            //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
            Document doc = new Document(MyDir + "Document.docx");
            // In Aspose.Words by default this option is set to true unlike Word
            TxtSaveOptions saveOptions = new TxtSaveOptions { AddBidiMarks = false };

            doc.Save(MyDir + @"\Artifacts\AddBidiMarks.txt", saveOptions);
            //ExEnd
        }
    }
}