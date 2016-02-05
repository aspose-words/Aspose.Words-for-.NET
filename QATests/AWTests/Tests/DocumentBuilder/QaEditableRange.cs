using System;
using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace QA_Tests.Tests
{
    /// <summary>
    /// Tests that verify editable regions of the document
    /// </summary>
    [TestFixture]
    class QaEditableRange : QaTestsBase
    {
        [Test]
        public void EditableRanges_AddEditableRanges()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Get paragraphs of the current document 
            Paragraph paraFirst = DocumentHelper.GetParagraph(doc, 0);
            Paragraph paraSecond = DocumentHelper.GetParagraph(doc, 1);

            builder.MoveTo(paraFirst);

            //Add EditableRangeStart to the first paragraph
            EditableRangeStart startRangeParaFirst = builder.StartEditableRange();

            builder.Writeln("EditableRange_1_1");
            builder.Writeln("EditableRange_1_2");

            //Mark the current position as an editable range end for "startRangeParaFirst"
            //"EndEditableRange()" closes the first created EditableRangeStart
            builder.EndEditableRange();

            //Add text to non-editable region of a document
            builder.Writeln("NotEditableRange_1_1");
            builder.Writeln("NotEditableRange_1_2");

            builder.MoveTo(paraSecond);

            //Add EditableRangeStart to the second paragraph
            EditableRangeStart startRangeParaSecond = builder.StartEditableRange();

            builder.Writeln("EditableRange_2_1");

            //Mark the current position as an editable range end for "startRangeParaSecond"
            //"EndEditableRange(EditableRangeStart)" closes EditableRangeStart which you specify in paramert
            builder.EndEditableRange(startRangeParaSecond);

            //Add text to non-editable region of a document
            builder.Writeln("NotEditableRange_2_1");

            //Sets the editor for editable range regions
            startRangeParaFirst.EditableRange.EditorGroup = EditorType.Everyone;
            startRangeParaSecond.EditableRange.EditorGroup = EditorType.Everyone;

            //Sets that the document read only and is password-protected
            doc.Protect(ProtectionType.ReadOnly, "123");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            NodeCollection startNodes = doc.GetChildNodes(NodeType.EditableRangeStart, true);

            //Assert that the document have nodes of EditableRangeStart
            Assert.AreEqual(2, startNodes.Count);

            //Assert that is the current region and structure is not broken
            Node startRangeRun1 = startNodes[0].NextSibling;
            Assert.AreEqual(startRangeRun1.GetText(), "EditableRange_1_1");

            //Assert that is the current region and structure is not broken
            Node startRangeRun2 = startNodes[1].NextSibling;
            Assert.AreEqual(startRangeRun2.GetText(), "EditableRange_2_1");

            //Assert that the document have nodes of EditableRangeEnd
            NodeCollection endNodes = doc.GetChildNodes(NodeType.EditableRangeEnd, true);
            Assert.AreEqual(2, endNodes.Count);

            //Assert that is the current region and structure is not broken
            Node endRangeRun1 = endNodes[0].NextSibling;
            Assert.AreEqual(endRangeRun1.GetText(), "NotEditableRange_1_1");

            //Assert that is the current region and structure is not broken
            Node endRangeRun2 = endNodes[1].NextSibling;
            Assert.AreEqual(endRangeRun2.GetText(), "NotEditableRange_2_1");
        }

        [Test]
        [ExpectedException(typeof(InvalidOperationException), ExpectedMessage = "EndEditableRange can not be called before StartEditableRange.")]
        public void EditableRanges_InvalidOperationException()
        {
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Is not valid structure for the current document
            builder.EndEditableRange();

            builder.StartEditableRange();
        }

        [Test]
        public void EditableRanges_WithoutEnd()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Add EditableRangeStart
            EditableRangeStart startRange1 = builder.StartEditableRange();

            builder.Writeln("EditableRange_1_1");
            builder.Writeln("EditableRange_1_2");

            //Sets the editor for editable range region
            startRange1.EditableRange.EditorGroup = EditorType.Everyone;

            //Sets that the document read only and is password-protected
            doc.Protect(ProtectionType.ReadOnly, "123");

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
