'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports NUnit.Framework
Imports Aspose.Words.Fields

Namespace Examples
	<TestFixture> _
	Public Class ExFormFields
		Inherits ExBase
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
			'ExSummary:Shows how to insert FormFields, set options annd gather them back in for use 
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Insert a text input field. The unique name of this field is "TextInput1", the other parameters define
			' what type of FormField it is, the format of the text, the field result and the maximum text length (0 = no limit)
			builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "", 0)
			'ExEnd
		End Sub
	End Class
End Namespace
