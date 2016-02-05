' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////



Imports Microsoft.VisualBasic
Imports System
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Aspose.Words
Imports Aspose.Words.Fields
Imports NUnit.Framework


Namespace ApiExamples.Fields
	<TestFixture> _
	Public Class ExField
		Inherits ApiExampleBase
		<Test> _
		Public Sub UpdateTOC()
			Dim doc As New Aspose.Words.Document()

			'ExStart
			'ExId:UpdateTOC
			'ExSummary:Shows how to completely rebuild TOC fields in the document by invoking field update.
			doc.UpdateFields()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetFieldType()
			Dim doc As New Aspose.Words.Document(MyDir & "Document.TableOfContents.doc")

			'ExStart
			'ExFor:FieldType
			'ExFor:FieldChar
			'ExFor:FieldChar.FieldType
			'ExSummary:Shows how to find the type of field that is represented by a node which is derived from FieldChar.
			Dim fieldStart As FieldChar = CType(doc.GetChild(NodeType.FieldStart, 0, True), FieldChar)
			Dim type As FieldType = fieldStart.FieldType
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetFieldFromDocument()
			'ExStart
			'ExFor:FieldChar.GetField
			'ExId:GetField
			'ExSummary:Demonstrates how to retrieve the field class from an existing FieldStart node in the document.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.TableOfContents.doc")

			Dim fieldStart As FieldStart = CType(doc.GetChild(NodeType.FieldStart, 0, True), FieldStart)

			' Retrieve the facade object which represents the field in the document.
			Dim field As Field = fieldStart.GetField()

			Console.WriteLine("Field code:" & field.GetFieldCode())
			Console.WriteLine("Field result: " & field.Result)
			Console.WriteLine("Is locked: " & field.IsLocked)

			' This updates only this field in the document.
			field.Update()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetFieldFromFieldCollection()
			'ExStart
			'ExId:GetFieldFromFieldCollection
			'ExSummary:Demonstrates how to retrieve a field using the range of a node.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.TableOfContents.doc")

			Dim field As Field = doc.Range.Fields(0)

			' This should be the first field in the document - a TOC field.
			Console.WriteLine(field.Type)
			'ExEnd
		End Sub

		<Test> _
		Public Sub InsertTCField()
			'ExStart
			'ExId:InsertTCField
			'ExSummary:Shows how to insert a TC field into the document using DocumentBuilder.
			' Create a blank document.
			Dim doc As New Aspose.Words.Document()

			' Create a document builder to insert content with.
			Dim builder As New DocumentBuilder(doc)

			' Insert a TC field at the current document builder position.
			builder.InsertField("TC ""Entry Text"" \f t")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ChangeLocale()
			' Create a blank document.
			Dim doc As New Aspose.Words.Document()
			Dim b As New DocumentBuilder(doc)
			b.InsertField("MERGEFIELD Date")

			'ExStart
			'ExId:ChangeCurrentCulture
			'ExSummary:Shows how to change the culture used in formatting fields during update.
			' Store the current culture so it can be set back once mail merge is complete.
			Dim currentCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
			' Set to German language so dates and numbers are formatted using this culture during mail merge.
			Thread.CurrentThread.CurrentCulture = New CultureInfo("de-DE")

			' Execute mail merge.
			doc.MailMerge.Execute(New String() { "Date" }, New Object() { DateTime.Now })

			' Restore the original culture.
			Thread.CurrentThread.CurrentCulture = currentCulture
			'ExEnd

			doc.Save(MyDir & "Field.ChangeLocale Out.doc")
		End Sub

		<Test> _
		Public Sub RemoveTOCFromDocument()
			'ExStart
			'ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
			'ExId:RemoveTableOfContents
			'ExSummary:Demonstrates how to remove a specified TOC from a document.
			' Open a document which contains a TOC.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.TableOfContents.doc")

			' Remove the first TOC from the document.
			Dim tocField As Field = doc.Range.Fields(0)
			tocField.Remove()

			' Save the output.
			doc.Save(MyDir & "Document.TableOfContentsRemoveTOC Out.doc")
			'ExEnd
		End Sub

		'ExStart
		'ExId:TCFieldsRangeReplace
		'ExSummary:Shows how to find and insert a TC field at text in a document. 
		<Test> _
		Public Sub InsertTCFieldsAtText()
			Dim doc As New Aspose.Words.Document()

			' Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
			doc.Range.Replace(New Regex("The Beginning"), New InsertTCFieldHandler("Chapter 1", "\l 1"), False)
		End Sub

		Public Class InsertTCFieldHandler
			Implements IReplacingCallback
			' Store the text and switches to be used for the TC fields.
			Private mFieldText As String
			Private mFieldSwitches As String

			''' <summary>
			''' The switches to use for each TC field. Can be an empty string or null.
			''' </summary>
			Public Sub New(ByVal switches As String)
				Me.New(String.Empty, switches)
				mFieldSwitches = switches
			End Sub

			''' <summary>
			''' The display text and switches to use for each TC field. Display name can be an empty string or null.
			''' </summary>
			Public Sub New(ByVal text As String, ByVal switches As String)
				mFieldText = text
				mFieldSwitches = switches
			End Sub

			Private Function IReplacingCallback_Replacing(ByVal args As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
				' Create a builder to insert the field.
				Dim builder As New DocumentBuilder(CType(args.MatchNode.Document, Aspose.Words.Document))
				' Move to the first node of the match.
				builder.MoveTo(args.MatchNode)

				' If the user specified text to be used in the field as display text then use that, otherwise use the 
				' match string as the display text.
				Dim insertText As String

				If (Not String.IsNullOrEmpty(mFieldText)) Then
					insertText = mFieldText
				Else
					insertText = args.Match.Value
				End If

				' Insert the TC field before this node using the specified string as the display text and user defined switches.
				builder.InsertField(String.Format("TC ""{0}"" {1}", insertText, mFieldSwitches))

				' We have done what we want so skip replacement.
				Return ReplaceAction.Skip
			End Function
		End Class
		'ExEnd
	End Class
End Namespace
