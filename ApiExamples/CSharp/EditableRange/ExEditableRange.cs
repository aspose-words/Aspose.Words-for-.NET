// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;



namespace ApiExamples.EditableRange
{
    [TestFixture]
    class ExEditableRange : ApiExampleBase
    {
        [Test]
        public void RemoveEx()
        {
            //ExStart
            //ExFor:EditableRange.Remove
            //ExSummary:Shows how to remove an editable range from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an EditableRange so we can remove it. Does not have to be well-formed.
            EditableRangeStart edRange1Start = builder.StartEditableRange();
            Aspose.Words.EditableRange editableRange1 = edRange1Start.EditableRange;
            builder.Writeln("Paragraph inside editable range");
            EditableRangeEnd edRange1End = builder.EndEditableRange();

            // Remove the range that was just made.
            editableRange1.Remove();
            //ExEnd
        }
    }
}
