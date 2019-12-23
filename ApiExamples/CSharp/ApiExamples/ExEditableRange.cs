// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Text;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExEditableRange : ApiExampleBase
    {
        [Test]
        public void RemovesEditableRange()
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

        //ExStart
        //ExFor:DocumentBuilder.StartEditableRange
        //ExFor:DocumentBuilder.EndEditableRange
        //ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
        //ExFor:EditableRange
        //ExFor:EditableRange.EditableRangeEnd
        //ExFor:EditableRange.EditableRangeStart
        //ExFor:EditableRange.Id
        //ExFor:EditableRange.SingleUser
        //ExFor:EditableRangeEnd
        //ExFor:EditableRangeEnd.Accept(DocumentVisitor)
        //ExFor:EditableRangeEnd.EditableRangeStart
        //ExFor:EditableRangeEnd.Id
        //ExFor:EditableRangeEnd.NodeType
        //ExFor:EditableRangeStart
        //ExFor:EditableRangeStart.Accept(DocumentVisitor)
        //ExFor:EditableRangeStart.EditableRange
        //ExFor:EditableRangeStart.Id
        //ExFor:EditableRangeStart.NodeType
        //ExSummary:Shows how to start and end an editable range.
        [Test] //ExSkip
        public void CreateEditableRanges()
        {
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start an editable range
            EditableRangeStart edRange1Start = builder.StartEditableRange();

            // An EditableRange object is created for the EditableRangeStart that we just made
            EditableRange editableRange1 = edRange1Start.EditableRange;

            // Put something inside the editable range
            builder.Writeln("Paragraph inside first editable range");

            // An editable range is well-formed if it has a start and an end
            // Multiple editable ranges can be nested and overlapping 
            EditableRangeEnd edRange1End = builder.EndEditableRange();

            // Explicitly state which EditableRangeStart a new EditableRangeEnd should be paired with
            EditableRangeStart edRange2Start = builder.StartEditableRange();
            builder.Writeln("Paragraph inside second editable range");
            EditableRange editableRange2 = edRange2Start.EditableRange;
            EditableRangeEnd edRange2End = builder.EndEditableRange(edRange2Start);

            // Editable range starts and ends have their own respective node types
            Assert.AreEqual(NodeType.EditableRangeStart, edRange1Start.NodeType);
            Assert.AreEqual(NodeType.EditableRangeEnd, edRange1End.NodeType);

            // Editable range IDs are unique and set automatically
            Assert.AreEqual(0, editableRange1.Id);
            Assert.AreEqual(1, editableRange2.Id);

            // Editable range starts and ends always belong to a range
            Assert.AreEqual(edRange1Start, editableRange1.EditableRangeStart);
            Assert.AreEqual(edRange1End, editableRange1.EditableRangeEnd);

            // They also inherit the ID of the entire editable range that they belong to
            Assert.AreEqual(editableRange1.Id, edRange1Start.Id);
            Assert.AreEqual(editableRange1.Id, edRange1End.Id);
            Assert.AreEqual(editableRange2.Id, edRange2Start.EditableRange.Id);
            Assert.AreEqual(editableRange2.Id, edRange2End.EditableRangeStart.EditableRange.Id);

            // If the editable range was found in a document, it will probably have something in the single user property
            // But if we make one programmatically, the property is empty by default
            Assert.AreEqual(string.Empty, editableRange1.SingleUser);

            // We have to set it ourselves if we want the ranges to belong to somebody
            editableRange1.SingleUser = "john.doe@myoffice.com";
            editableRange2.SingleUser = "jane.doe@myoffice.com";

            // Initialize a custom visitor for editable ranges that will print their contents 
            EditableRangeInfoPrinter editableRangeReader = new EditableRangeInfoPrinter();

            // Both the start and end of an editable range can accept visitors, but not the editable range itself
            edRange1Start.Accept(editableRangeReader);
            edRange2End.Accept(editableRangeReader);

            // Or, if we want to go over all the editable ranges in a document, we can get the document to accept the visitor
            editableRangeReader.Reset();
            doc.Accept(editableRangeReader);

            Console.WriteLine(editableRangeReader.ToText());
        }

        /// <summary>
        /// Visitor implementation that prints attributes and contents of ranges.
        /// </summary>
        public class EditableRangeInfoPrinter : DocumentVisitor
        {
            public EditableRangeInfoPrinter()
            {
                mBuilder = new StringBuilder();
            }

            public string ToText()
            {
                return mBuilder.ToString();
            }

            public void Reset()
            {
                mBuilder.Clear();
                mInsideEditableRange = false;
            }

            /// <summary>
            /// Called when an EditableRangeStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitEditableRangeStart(EditableRangeStart editableRangeStart)
            {
                mBuilder.AppendLine(" -- Editable range found! -- ");
                mBuilder.AppendLine("\tID: " + editableRangeStart.Id);
                mBuilder.AppendLine("\tUser: " + editableRangeStart.EditableRange.SingleUser);
                mBuilder.AppendLine("\tContents: ");

                mInsideEditableRange = true;

                // Let the visitor continue visiting other nodes
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an EditableRangeEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
            {
                mBuilder.AppendLine(" -- End of editable range -- ");

                mInsideEditableRange = false;

                // Let the visitor continue visiting other nodes
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Run node is encountered in the document. Only runs within editable ranges have their contents recorded.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mInsideEditableRange) mBuilder.AppendLine("\t\"" + run.Text + "\"");

                // Let the visitor continue visiting other nodes
                return VisitorAction.Continue;
            }

            private bool mInsideEditableRange;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        [Test]
        public void IncorrectStructureException()
        {
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Checking that isn't valid structure for the current document
            Assert.That(() => builder.EndEditableRange(), Throws.TypeOf<InvalidOperationException>());

            builder.StartEditableRange();
        }

        [Test]
        public void IncorrectStructureDoNotAdded()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //ExStart
            //ExFor:EditableRange.EditorGroup
            //ExFor:EditorType
            //ExSummary:Shows how to add editing group for editable ranges
            EditableRangeStart startRange1 = builder.StartEditableRange();

            builder.Writeln("EditableRange_1_1");
            builder.Writeln("EditableRange_1_2");

            // Sets the editor for editable range region
            startRange1.EditableRange.EditorGroup = EditorType.Everyone;
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            // Assert that it's not valid structure and editable ranges aren't added to the current document
            NodeCollection startNodes = doc.GetChildNodes(NodeType.EditableRangeStart, true);
            Assert.AreEqual(0, startNodes.Count);

            NodeCollection endNodes = doc.GetChildNodes(NodeType.EditableRangeEnd, true);
            Assert.AreEqual(0, endNodes.Count);
        }
    }
}