' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.IO

Imports Aspose.Words

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Friend Class ExEditableRange
		Inherits ApiExampleBase
		<Test> _
		Public Sub RemoveEx()
			'ExStart
			'ExFor:EditableRange.Remove
			'ExSummary:Shows how to remove an editable range from a document.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim builder As New DocumentBuilder(doc)

			' Create an EditableRange so we can remove it. Does not have to be well-formed.
			Dim edRange1Start As EditableRangeStart = builder.StartEditableRange()
			Dim editableRange1 As EditableRange = edRange1Start.EditableRange
			builder.Writeln("Paragraph inside editable range")
			Dim edRange1End As EditableRangeEnd = builder.EndEditableRange()

			' Remove the range that was just made.
			editableRange1.Remove()
			'ExEnd
		End Sub

		<Test> _
		Public Sub EditableRangeEx()
			'ExStart
			'ExFor:DocumentBuilder.StartEditableRange
			'ExFor:DocumentBuilder.EndEditableRange()
			'ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
			'ExSummary:Shows how to start and end an editable range.
			Dim doc As New Document(MyDir & "Document.doc")
			Dim builder As New DocumentBuilder(doc)

			' Start an editable range.
			Dim edRange1Start As EditableRangeStart = builder.StartEditableRange()

			' An EditableRange object is created for the EditableRangeStart that we just made.
			Dim editableRange1 As EditableRange = edRange1Start.EditableRange

			' Put something inside the editable range.
			builder.Writeln("Paragraph inside first editable range")

			' An editable range is well-formed if it has a start and an end. 
			' Multiple editable ranges can be nested and overlapping. 
			Dim edRange1End As EditableRangeEnd = builder.EndEditableRange()

			' Both the start and end automatically belong to editableRange1.
			Console.WriteLine(editableRange1.EditableRangeStart.Equals(edRange1Start)) ' True
			Console.WriteLine(editableRange1.EditableRangeEnd.Equals(edRange1End)) ' True

			' Explicitly state which EditableRangeStart a new EditableRangeEnd should be paired with.
			Dim edRange2Start As EditableRangeStart = builder.StartEditableRange()
			builder.Writeln("Paragraph inside second editable range")
			Dim editableRange2 As EditableRange = edRange2Start.EditableRange
			Dim edRange2End As EditableRangeEnd = builder.EndEditableRange(edRange2Start)

			' Both the start and end automatically belong to editableRange2.
			Console.WriteLine(editableRange2.EditableRangeStart.Equals(edRange2Start)) ' True
			Console.WriteLine(editableRange2.EditableRangeEnd.Equals(edRange2End)) ' True
			'ExEnd
		End Sub

		<Test, ExpectedException(GetType(InvalidOperationException), ExpectedMessage := "EndEditableRange can not be called before StartEditableRange.")> _
		Public Sub IncorrectStructureException()
			Dim doc As New Document()

			Dim builder As New DocumentBuilder(doc)

			'Is not valid structure for the current document
			builder.EndEditableRange()

			builder.StartEditableRange()
		End Sub

		<Test> _
		Public Sub IncorrectStructureDoNotAdded()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			Dim builder As New DocumentBuilder(doc)

			'ExStart
			'ExFor: EditableRange.EditorGroup
			'ExSummary:Shows how to add editing group for editableranges
			'Add EditableRangeStart
			Dim startRange1 As EditableRangeStart = builder.StartEditableRange()

			builder.Writeln("EditableRange_1_1")
			builder.Writeln("EditableRange_1_2")

			'Sets the editor for editable range region
			startRange1.EditableRange.EditorGroup = EditorType.Everyone
			'ExEnd

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			'Assert that it's not valid structure and editable ranges aren't added to the current document
			Dim startNodes As NodeCollection = doc.GetChildNodes(NodeType.EditableRangeStart, True)
			Assert.AreEqual(0, startNodes.Count)

			Dim endNodes As NodeCollection = doc.GetChildNodes(NodeType.EditableRangeEnd, True)
			Assert.AreEqual(0, endNodes.Count)
		End Sub
	End Class
End Namespace
