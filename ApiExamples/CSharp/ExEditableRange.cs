// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;

using Aspose.Words;

using NUnit.Framework;

namespace ApiExamples
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
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create an EditableRange so we can remove it. Does not have to be well-formed.
            EditableRangeStart edRange1Start = builder.StartEditableRange();
            EditableRange editableRange1 = edRange1Start.EditableRange;
            builder.Writeln("Paragraph inside editable range");
            EditableRangeEnd edRange1End = builder.EndEditableRange();

            // Remove the range that was just made.
            editableRange1.Remove();
            //ExEnd
        }

        [Test]
        public void EditableRangeEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.StartEditableRange
            //ExFor:DocumentBuilder.EndEditableRange()
            //ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
            //ExSummary:Shows how to start and end an editable range.
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start an editable range.
            EditableRangeStart edRange1Start = builder.StartEditableRange();

            // An EditableRange object is created for the EditableRangeStart that we just made.
            EditableRange editableRange1 = edRange1Start.EditableRange;

            // Put something inside the editable range.
            builder.Writeln("Paragraph inside first editable range");

            // An editable range is well-formed if it has a start and an end. 
            // Multiple editable ranges can be nested and overlapping. 
            EditableRangeEnd edRange1End = builder.EndEditableRange();

            // Both the start and end automatically belong to editableRange1.
            Console.WriteLine(editableRange1.EditableRangeStart.Equals(edRange1Start)); // True
            Console.WriteLine(editableRange1.EditableRangeEnd.Equals(edRange1End)); // True

            // Explicitly state which EditableRangeStart a new EditableRangeEnd should be paired with.
            EditableRangeStart edRange2Start = builder.StartEditableRange();
            builder.Writeln("Paragraph inside second editable range");
            EditableRange editableRange2 = edRange2Start.EditableRange;
            EditableRangeEnd edRange2End = builder.EndEditableRange(edRange2Start);

            // Both the start and end automatically belong to editableRange2.
            Console.WriteLine(editableRange2.EditableRangeStart.Equals(edRange2Start)); // True
            Console.WriteLine(editableRange2.EditableRangeEnd.Equals(edRange2End)); // True
            //ExEnd
        }

        [Test]
        public void IncorrectStructureException()
        {
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Is not valid structure for the current document
            Assert.That(() => builder.EndEditableRange(), Throws.TypeOf<InvalidOperationException>());

            builder.StartEditableRange();
        }

        [Test]
        public void IncorrectStructureDoNotAdded()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            DocumentBuilder builder = new DocumentBuilder(doc);

            //ExStart
            //ExFor: EditableRange.EditorGroup
            //ExSummary:Shows how to add editing group for editableranges
            //Add EditableRangeStart
            EditableRangeStart startRange1 = builder.StartEditableRange();

            builder.Writeln("EditableRange_1_1");
            builder.Writeln("EditableRange_1_2");

            //Sets the editor for editable range region
            startRange1.EditableRange.EditorGroup = EditorType.Everyone;
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            //Assert that it's not valid structure and editable ranges aren't added to the current document
            NodeCollection startNodes = doc.GetChildNodes(NodeType.EditableRangeStart, true);
            Assert.AreEqual(0, startNodes.Count);

            NodeCollection endNodes = doc.GetChildNodes(NodeType.EditableRangeEnd, true);
            Assert.AreEqual(0, endNodes.Count);
        }
    }
}
