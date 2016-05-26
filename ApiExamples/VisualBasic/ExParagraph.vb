Imports Microsoft.VisualBasic
Imports System

Imports Aspose.Words
Imports Aspose.Words.Fields

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Friend Class ExParagraph
		Inherits ApiExampleBase
		<Test> _
		Public Sub InsertField()
			'ExStart
			'ExFor:Paragraph.InsertField
			'ExSummary:Shows how to insert field using several methods: "field code", "field code and field value", "field code and field value after a run of text"
			Dim doc As New Document()

			'Get the first paragraph of the document
			Dim para As Paragraph = doc.FirstSection.Body.FirstParagraph

			'Inseting field using field code
			'Note: All methods support inserting field after some node. Just set "true" in the "isAfter" parameter
			para.InsertField(" AUTHOR ", Nothing, False)

			'Using field type
			'Note:
			'1. For inserting field using field type, you can choose, update field before or after you open the document ("updateField" parameter)
			'2. For other methods it's works automatically
			para.InsertField(FieldType.FieldAuthor, False, Nothing, True)

			'Using field code and field value
			para.InsertField(" AUTHOR ", "Test Field Value", Nothing, False)

			'Add a run of text
			Dim run As New Run(doc) With {.Text = " Hello World!"}
			para.AppendChild(run)

			'Using field code and field value before a run of text
			'Note: For inserting field before/after a run of text you can use all methods above, just add ref on your text ("refNode" parameter)
			para.InsertField(" AUTHOR ", "Test Field Value", run, False)
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertFieldBeforeTextInParagraph()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			InsertFieldUsingFieldCode(doc, " AUTHOR ", Nothing, False, 1)

			Assert.AreEqual(ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & "Test Author" & ChrW(&H0015).ToString() & "Hello World!" & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		<Test> _
		Public Sub InsertFieldAfterTextInParagraph()
			Dim [date] As String = DateTime.Today.ToString("d")

			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			InsertFieldUsingFieldCode(doc, " DATE ", Nothing, True, 1)

			Assert.AreEqual(String.Format("Hello World!" & ChrW(&H0013).ToString() & " DATE " & ChrW(&H0014).ToString() & "{0}" & ChrW(&H0015).ToString() & Constants.vbCr, [date]), DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		<Test> _
		Public Sub InsertFieldBeforeTextInParagraphWithoutUpdateField()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, False, Nothing, False, 1)

			Assert.AreEqual(ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & ChrW(&H0015).ToString() & "Hello World!" & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		<Test> _
		Public Sub InsertFieldAfterTextInParagraphWithoutUpdateField()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, False, Nothing, True, 1)

			Assert.AreEqual("Hello World!" & ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & ChrW(&H0015).ToString() & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		<Test> _
		Public Sub InsertFieldWithoutSeparator()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			InsertFieldUsingFieldType(doc, FieldType.FieldListNum, True, Nothing, False, 1)

			Assert.AreEqual(ChrW(&H0013).ToString() & " LISTNUM " & ChrW(&H0015).ToString() & "Hello World!" & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		<Test> _
		Public Sub InsertFieldBeforeParagraphWithoutDocumentAuthor()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()
			doc.BuiltInDocumentProperties.Author = ""

			InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", Nothing, Nothing, False, 1)

			Assert.AreEqual(ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & ChrW(&H0015).ToString() & "Hello World!" & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		<Test> _
		Public Sub InsertFieldAfterParagraphWithoutChangingDocumentAuthor()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", Nothing, Nothing, True, 1)

			Assert.AreEqual("Hello World!" & ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & ChrW(&H0015).ToString() & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		<Test> _
		Public Sub InsertFieldBeforeRunText()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			'Add some text into the paragraph
			Dim run As Run = DocumentHelper.InsertNewRun(doc, " Hello World!", 1)

			InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, False, 1)

			Assert.AreEqual("Hello World!" & ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & "Test Field Value" & ChrW(&H0015).ToString() & " Hello World!" & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		<Test> _
		Public Sub InsertFieldAfterRunText()
			Dim doc As Document = DocumentHelper.CreateDocumentFillWithDummyText()

			'Add some text into the paragraph
			Dim run As Run = DocumentHelper.InsertNewRun(doc, " Hello World!", 1)

			InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, True, 1)

			Assert.AreEqual("Hello World! Hello World!" & ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & ChrW(&H0015).ToString() & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		''' <summary>
		''' Test for WORDSNET-12396
		''' </summary>
		<Test> _
		Public Sub InsertFieldEmptyParagraphWithoutUpdateField()
			Dim doc As Document = DocumentHelper.CreateDocumentWithoutDummyText()

			InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, False, Nothing, False, 1)

			Assert.AreEqual(ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & ChrW(&H0015).ToString() & Constants.vbFormFeed, DocumentHelper.GetParagraphText(doc, 1))
		End Sub

		''' <summary>
		''' Test for WORDSNET-12397
		''' </summary>
		<Test> _
		Public Sub InsertFieldEmptyParagraphWithUpdateField()
			Dim doc As Document = DocumentHelper.CreateDocumentWithoutDummyText()

			InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, True, Nothing, False, 0)

			Assert.AreEqual(ChrW(&H0013).ToString() & " AUTHOR " & ChrW(&H0014).ToString() & "Test Author" & ChrW(&H0015).ToString() & Constants.vbCr, DocumentHelper.GetParagraphText(doc, 0))
		End Sub

		<Test> _
		Public Sub GetFormatRevision()
			'ExStart
			'ExFor:Paragraph.IsFormatRevision
			'ExSummary:Shows how to get information about whether this object was formatted in Microsoft Word while change tracking was enabled
			Dim doc As New Document(MyDir & "Paragraph.IsFormatRevision.docx")

			Dim firstParagraph As Paragraph = DocumentHelper.GetParagraph(doc, 0)
			Assert.IsTrue(firstParagraph.IsFormatRevision)
			'ExEnd

			Dim secondParagraph As Paragraph = DocumentHelper.GetParagraph(doc, 1)
			Assert.IsFalse(secondParagraph.IsFormatRevision)
		End Sub

		''' <summary>
		''' Insert field into the first paragraph of the current document using field type
		''' </summary>
		Private Shared Sub InsertFieldUsingFieldType(ByVal doc As Document, ByVal fieldType As FieldType, ByVal updateField As Boolean, ByVal refNode As Node, ByVal isAfter As Boolean, ByVal paraIndex As Integer)
			Dim para As Paragraph = DocumentHelper.GetParagraph(doc, paraIndex)
			para.InsertField(fieldType, updateField, refNode, isAfter)
		End Sub

		''' <summary>
		''' Insert field into the first paragraph of the current document using field code
		''' </summary>
		Private Shared Sub InsertFieldUsingFieldCode(ByVal doc As Document, ByVal fieldCode As String, ByVal refNode As Node, ByVal isAfter As Boolean, ByVal paraIndex As Integer)
			Dim para As Paragraph = DocumentHelper.GetParagraph(doc, paraIndex)
			para.InsertField(fieldCode, refNode, isAfter)
		End Sub

		''' <summary>
		''' Insert field into the first paragraph of the current document using field code and field string
		''' </summary>
		Private Shared Sub InsertFieldUsingFieldCodeFieldString(ByVal doc As Document, ByVal fieldCode As String, ByVal fieldValue As String, ByVal refNode As Node, ByVal isAfter As Boolean, ByVal paraIndex As Integer)
			Dim para As Paragraph = DocumentHelper.GetParagraph(doc, paraIndex)
			para.InsertField(fieldCode, fieldValue, refNode, isAfter)
		End Sub
	End Class
End Namespace
