' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports Aspose.Words
Imports Aspose.Words.Fields

Imports NUnit.Framework

	Imports System.IO
Namespace ApiExamples

	<TestFixture> _
	Public Class ExFormFields
		Inherits ApiExampleBase
		<Test> _
		Public Sub FormFieldsGetFormFieldsCollection()
			'ExStart
			'ExFor:Range.FormFields
			'ExFor:FormFieldCollection
			'ExId:FormFieldsGetFormFieldsCollection
			'ExSummary:Shows how to get a collection of form fields.
			Dim doc As New Document(MyDir & "FormFields.doc")
			Dim formFields As FormFieldCollection = doc.Range.FormFields
			'ExEnd
		End Sub

		<Test> _
		Public Sub FormFieldsGetByName()
			'ExStart
			'ExFor:FormField
			'ExId:FormFieldsGetByName
			'ExSummary:Shows how to access form fields.
			Dim doc As New Document(MyDir & "FormFields.doc")
			Dim documentFormFields As FormFieldCollection = doc.Range.FormFields

			Dim formField1 As FormField = documentFormFields(3)
			Dim formField2 As FormField = documentFormFields("CustomerName")
			'ExEnd
		End Sub

		<Test> _
		Public Sub FormFieldsWorkWithProperties()
			'ExStart
			'ExFor:FormField
			'ExFor:FormField.Result
			'ExFor:FormField.Type
			'ExFor:FormField.Name
			'ExId:FormFieldsWorkWithProperties
			'ExSummary:Shows how to work with form field name, type, and result.
			Dim doc As New Document(MyDir & "FormFields.doc")

			Dim formField As FormField = doc.Range.FormFields(3)

			If formField.Type.Equals(FieldType.FieldFormTextInput) Then
				formField.Result = "My name is " & formField.Name
			End If
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertAndRetrieveFormFields()
			'ExStart
			'ExFor:DocumentBuilder.InsertTextInput
			'ExId:FormFieldsInsertAndRetrieve
			'ExSummary:Shows how to insert form fields, set options and gather them back in for use 
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
			' what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
			builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "", 0)
			'ExEnd
		End Sub

		<Test> _
		Public Sub DeleteFormField()
			'ExStart
			'ExFor:FormField.RemoveField
			'ExSummary:Shows how to delete complete form field
			Dim doc As New Document(MyDir & "FormFields.doc")

			Dim formField As FormField = doc.Range.FormFields(3)
			formField.RemoveField()
			'ExEnd

			Dim formFieldAfter As FormField = doc.Range.FormFields(3)

			Assert.IsNull(formFieldAfter)
		End Sub

		<Test> _
		Public Sub DeleteFormFieldAssociatedWithTheFormField()
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			builder.StartBookmark("MyBookmark")
			builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "TestFormField", "SomeText", 0)
			builder.EndBookmark("MyBookmark")

			Dim dstStream As New MemoryStream()
			doc.Save(dstStream, SaveFormat.Docx)

			Dim bookmarkBeforeDeleteFormField As BookmarkCollection = doc.Range.Bookmarks
			Assert.AreEqual("MyBookmark", bookmarkBeforeDeleteFormField(0).Name)

			Dim formField As FormField = doc.Range.FormFields(0)
			formField.RemoveField()

			Dim bookmarkAfterDeleteFormField As BookmarkCollection = doc.Range.Bookmarks
			Assert.AreEqual("MyBookmark", bookmarkAfterDeleteFormField(0).Name)
		End Sub
	End Class
End Namespace
