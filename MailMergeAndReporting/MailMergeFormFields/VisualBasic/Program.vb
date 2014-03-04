' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
'ExStart
'ExId:MailMergeFormFields
'ExSummary:Complete source code of a program that inserts checkboxes and text input form fields into a document during mail merge.

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports Aspose.Words
Imports Aspose.Words.Fields
Imports Aspose.Words.Reporting

Namespace MailMergeFormFields
	''' <summary>
	''' This sample shows how to insert check boxes and text input form fields during mail merge into a document.
	''' </summary>
	Public Class Program
		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		Public Shared Sub Main(ByVal args() As String)
			Dim program As New Program()
			program.Execute()
		End Sub

		Private Sub Execute()
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Load the template document.
			Dim doc As New Document(dataDir & "Template.doc")

			' Setup mail merge event handler to do the custom work.
			doc.MailMerge.FieldMergingCallback = New HandleMergeField()

			' This is the data for mail merge.
			Dim fieldNames() As String = {"RecipientName", "SenderName", "FaxNumber", "PhoneNumber", "Subject", "Body", "Urgent", "ForReview", "PleaseComment"}
			Dim fieldValues() As Object = {"Josh", "Jenny", "123456789", "", "Hello", "Test message 1", True, False, True}

			' Execute the mail merge.
			doc.MailMerge.Execute(fieldNames, fieldValues)

			' Save the finished document.
			doc.Save(dataDir & "Template Out.doc")
		End Sub

		Private Class HandleMergeField
			Implements IFieldMergingCallback
			''' <summary>
			''' This handler is called for every mail merge field found in the document,
			'''  for every record found in the data source.
			''' </summary>
			Private Sub IFieldMergingCallback_FieldMerging(ByVal e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
				If mBuilder Is Nothing Then
					mBuilder = New DocumentBuilder(e.Document)
				End If

				' We decided that we want all boolean values to be output as check box form fields.
				If TypeOf e.FieldValue Is Boolean Then
					' Move the "cursor" to the current merge field.
					mBuilder.MoveToMergeField(e.FieldName)

					' It is nice to give names to check boxes. Lets generate a name such as MyField21 or so.
					Dim checkBoxName As String = String.Format("{0}{1}", e.FieldName, e.RecordIndex)

					' Insert a check box.
					mBuilder.InsertCheckBox(checkBoxName, CBool(e.FieldValue), 0)

					' Nothing else to do for this field.
					Return
				End If

				' Another example, we want the Subject field to come out as text input form field.
				If e.FieldName = "Subject" Then
					mBuilder.MoveToMergeField(e.FieldName)
					Dim textInputName As String = String.Format("{0}{1}", e.FieldName, e.RecordIndex)
					mBuilder.InsertTextInput(textInputName, TextFormFieldType.Regular, "", CStr(e.FieldValue), 0)
				End If
			End Sub

			Private Sub ImageFieldMerging(ByVal args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
				' Do nothing.
			End Sub

			Private mBuilder As DocumentBuilder
		End Class
	End Class
End Namespace
'ExEnd