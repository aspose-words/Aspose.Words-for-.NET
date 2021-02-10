// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Text;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExEditableRange : ApiExampleBase
    {
        [Test]
        public void CreateAndRemove()
        {
            //ExStart
            //ExFor:DocumentBuilder.StartEditableRange
            //ExFor:DocumentBuilder.EndEditableRange
            //ExFor:EditableRange
            //ExFor:EditableRange.EditableRangeEnd
            //ExFor:EditableRange.EditableRangeStart
            //ExFor:EditableRange.Id
            //ExFor:EditableRange.Remove
            //ExFor:EditableRangeEnd.EditableRangeStart
            //ExFor:EditableRangeEnd.Id
            //ExFor:EditableRangeEnd.NodeType
            //ExFor:EditableRangeStart.EditableRange
            //ExFor:EditableRangeStart.Id
            //ExFor:EditableRangeStart.NodeType
            //ExSummary:Shows how to work with an editable range.
            Document doc = new Document();
            doc.Protect(ProtectionType.ReadOnly, "MyPassword");

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! Since we have set the document's protection level to read-only," +
                            " we cannot edit this paragraph without the password.");

            // Editable ranges allow us to leave parts of protected documents open for editing.
            EditableRangeStart editableRangeStart = builder.StartEditableRange();
            builder.Writeln("This paragraph is inside an editable range, and can be edited.");
            EditableRangeEnd editableRangeEnd = builder.EndEditableRange();

            // A well-formed editable range has a start node, and end node.
            // These nodes have matching IDs and encompass editable nodes.
            EditableRange editableRange = editableRangeStart.EditableRange;

            Assert.AreEqual(editableRangeStart.Id, editableRange.Id);
            Assert.AreEqual(editableRangeEnd.Id, editableRange.Id);
            
            // Different parts of the editable range link to each other.
            Assert.AreEqual(editableRangeStart.Id, editableRange.EditableRangeStart.Id);
            Assert.AreEqual(editableRangeStart.Id, editableRangeEnd.EditableRangeStart.Id);
            Assert.AreEqual(editableRange.Id, editableRangeStart.EditableRange.Id);
            Assert.AreEqual(editableRangeEnd.Id, editableRange.EditableRangeEnd.Id);

            // We can access the node types of each part like this. The editable range itself is not a node,
            // but an entity which consists of a start, an end, and their enclosed contents.
            Assert.AreEqual(NodeType.EditableRangeStart, editableRangeStart.NodeType);
            Assert.AreEqual(NodeType.EditableRangeEnd, editableRangeEnd.NodeType);

            builder.Writeln("This paragraph is outside the editable range, and cannot be edited.");

            doc.Save(ArtifactsDir + "EditableRange.CreateAndRemove.docx");

            // Remove an editable range. All the nodes that were inside the range will remain intact.
            editableRange.Remove();
            //ExEnd

            Assert.AreEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                            "This paragraph is inside an editable range, and can be edited.\r" +
                            "This paragraph is outside the editable range, and cannot be edited.", doc.GetText().Trim());
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.EditableRangeStart, true).Count);

            doc = new Document(ArtifactsDir + "EditableRange.CreateAndRemove.docx");

            Assert.AreEqual(ProtectionType.ReadOnly, doc.ProtectionType);
            Assert.AreEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                            "This paragraph is inside an editable range, and can be edited.\r" +
                            "This paragraph is outside the editable range, and cannot be edited.", doc.GetText().Trim());

            editableRange = ((EditableRangeStart)doc.GetChild(NodeType.EditableRangeStart, 0, true)).EditableRange;

            TestUtil.VerifyEditableRange(0, string.Empty, EditorType.Unspecified, editableRange);
        }


        [Test]
        public void Nested()
        {
            //ExStart
            //ExFor:DocumentBuilder.StartEditableRange
            //ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
            //ExFor:EditableRange.EditorGroup
            //ExSummary:Shows how to create nested editable ranges.
            Document doc = new Document();
            doc.Protect(ProtectionType.ReadOnly, "MyPassword");

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! Since we have set the document's protection level to read-only, " +
                            "we cannot edit this paragraph without the password.");
             
            // Create two nested editable ranges.
            EditableRangeStart outerEditableRangeStart = builder.StartEditableRange();
            builder.Writeln("This paragraph inside the outer editable range and can be edited.");

            EditableRangeStart innerEditableRangeStart = builder.StartEditableRange();
            builder.Writeln("This paragraph inside both the outer and inner editable ranges and can be edited.");

            // Currently, the document builder's node insertion cursor is in more than one ongoing editable range.
            // When we want to end an editable range in this situation,
            // we need to specify which of the ranges we wish to end by passing its EditableRangeStart node.
            builder.EndEditableRange(innerEditableRangeStart);

            builder.Writeln("This paragraph inside the outer editable range and can be edited.");

            builder.EndEditableRange(outerEditableRangeStart);

            builder.Writeln("This paragraph is outside any editable ranges, and cannot be edited.");

            // If a region of text has two overlapping editable ranges with specified groups,
            // the combined group of users excluded by both groups are prevented from editing it.
            outerEditableRangeStart.EditableRange.EditorGroup = EditorType.Everyone;
            innerEditableRangeStart.EditableRange.EditorGroup = EditorType.Contributors;

            doc.Save(ArtifactsDir + "EditableRange.Nested.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "EditableRange.Nested.docx");

            Assert.AreEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                            "This paragraph inside the outer editable range and can be edited.\r" +
                            "This paragraph inside both the outer and inner editable ranges and can be edited.\r" +
                            "This paragraph inside the outer editable range and can be edited.\r" +
                            "This paragraph is outside any editable ranges, and cannot be edited.", doc.GetText().Trim());

            EditableRange editableRange = ((EditableRangeStart)doc.GetChild(NodeType.EditableRangeStart, 0, true)).EditableRange;

            TestUtil.VerifyEditableRange(0, string.Empty, EditorType.Everyone, editableRange);

            editableRange = ((EditableRangeStart)doc.GetChild(NodeType.EditableRangeStart, 1, true)).EditableRange;

            TestUtil.VerifyEditableRange(1, string.Empty, EditorType.Contributors, editableRange);
        }

        //ExStart

        //ExFor:EditableRange
        //ExFor:EditableRange.EditorGroup
        //ExFor:EditableRange.SingleUser
        //ExFor:EditableRangeEnd
        //ExFor:EditableRangeEnd.Accept(DocumentVisitor)
        //ExFor:EditableRangeStart
        //ExFor:EditableRangeStart.Accept(DocumentVisitor)
        //ExFor:EditorType
        //ExSummary:Shows how to limit the editing rights of editable ranges to a specific group/user.
        [Test] //ExSkip
        public void Visitor()
        {
            Document doc = new Document();
            doc.Protect(ProtectionType.ReadOnly, "MyPassword");

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! Since we have set the document's protection level to read-only," +
                            " we cannot edit this paragraph without the password.");

            // When we write-protect documents, editable ranges allow us to pick specific areas that users may edit.
            // There are two mutually exclusive ways to narrow down the list of allowed editors.
            // 1 -  Specify a user:
            EditableRange editableRange = builder.StartEditableRange().EditableRange;
            editableRange.SingleUser = "john.doe@myoffice.com";
            builder.Writeln($"This paragraph is inside the first editable range, can only be edited by {editableRange.SingleUser}.");
            builder.EndEditableRange();

            Assert.AreEqual(EditorType.Unspecified, editableRange.EditorGroup);

            // 2 -  Specify a group that allowed users are associated with:
            editableRange = builder.StartEditableRange().EditableRange;
            editableRange.EditorGroup = EditorType.Administrators;
            builder.Writeln($"This paragraph is inside the first editable range, can only be edited by {editableRange.EditorGroup}.");
            builder.EndEditableRange();

            Assert.AreEqual(string.Empty, editableRange.SingleUser);

            builder.Writeln("This paragraph is outside the editable range, and cannot be edited by anybody.");

            // Print details and contents of every editable range in the document.
            EditableRangePrinter editableRangePrinter = new EditableRangePrinter();

            doc.Accept(editableRangePrinter);

            Console.WriteLine(editableRangePrinter.ToText());
        }

        /// <summary>
        /// Collects properties and contents of visited editable ranges in a string.
        /// </summary>
        public class EditableRangePrinter : DocumentVisitor
        {
            public EditableRangePrinter()
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
                mBuilder.AppendLine("\tID:\t\t" + editableRangeStart.Id);
                if (editableRangeStart.EditableRange.SingleUser == string.Empty)
                    mBuilder.AppendLine("\tGroup:\t" + editableRangeStart.EditableRange.EditorGroup);
                else
                    mBuilder.AppendLine("\tUser:\t" + editableRangeStart.EditableRange.SingleUser);
                mBuilder.AppendLine("\tContents:");

                mInsideEditableRange = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an EditableRangeEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
            {
                mBuilder.AppendLine(" -- End of editable range --\n");

                mInsideEditableRange = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Run node is encountered in the document. This visitor only records runs that are inside editable ranges.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mInsideEditableRange) mBuilder.AppendLine("\t\"" + run.Text + "\"");

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

            // Assert that isn't valid structure for the current document.
            Assert.That(() => builder.EndEditableRange(), Throws.TypeOf<InvalidOperationException>());

            builder.StartEditableRange();
        }

        [Test]
        public void IncorrectStructureDoNotAdded()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();
            DocumentBuilder builder = new DocumentBuilder(doc);

            EditableRangeStart startRange1 = builder.StartEditableRange();

            builder.Writeln("EditableRange_1_1");
            builder.Writeln("EditableRange_1_2");

            startRange1.EditableRange.EditorGroup = EditorType.Everyone;
            doc = DocumentHelper.SaveOpen(doc);

            // Assert that it's not valid structure and editable ranges aren't added to the current document.
            NodeCollection startNodes = doc.GetChildNodes(NodeType.EditableRangeStart, true);
            Assert.AreEqual(0, startNodes.Count);

            NodeCollection endNodes = doc.GetChildNodes(NodeType.EditableRangeEnd, true);
            Assert.AreEqual(0, endNodes.Count);
        }
    }
}