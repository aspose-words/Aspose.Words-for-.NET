'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////



Imports Microsoft.VisualBasic
Imports NUnit.Framework
Imports System
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.Fields
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Threading

Namespace Examples
	<TestFixture> _
	Public Class ExField
		Inherits ExBase
		<Test> _
		Public Sub UpdateTOC()
			Dim doc As New Document()

			'ExStart
			'ExId:UpdateTOC
			'ExSummary:Shows how to completely rebuild TOC fields in the document by invoking field update.
			doc.UpdateFields()
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetFieldType()
			Dim doc As New Document(MyDir & "Document.TableOfContents.doc")

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
		Public Sub InsertTCField()
			'ExStart
			'ExId:InsertTCField
			'ExSummary:Shows how to insert a TC field into the document using DocumentBuilder.
			' Create a blank document.
			Dim doc As New Document()

			' Create a document builder to insert content with.
			Dim builder As New DocumentBuilder(doc)

			' Insert a TC field at the current document builder position.
			builder.InsertField("TC ""Entry Text"" \f t")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ChangeLocale()
			' Create a blank document.
			Dim doc As New Document()
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

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub RemoveTOCFromDocumentCaller()
			RemoveTOCFromDocument()
		End Sub

		'ExStart
		'ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
		'ExId:RemoveTableOfContents
		'ExSummary:Demonstrates how to remove a specified TOC from a document.
		Public Sub RemoveTOCFromDocument()
			' Open a document which contains a TOC.
			Dim doc As New Document(MyDir & "Document.TableOfContents.doc")

			' Remove the first table of contents from the document.
			RemoveTableOfContents(doc, 0)

			' Save the output.
			doc.Save(MyDir & "Document.TableOfContentsRemoveTOC Out.doc")
		End Sub

		''' <summary>
		''' Removes the specified table of contents field from the document.
		''' </summary>
		''' <param name="doc">The document to remove the field from.</param>
		''' <param name="index">The zero-based index of the TOC to remove.</param>
		Private Shared Sub RemoveTableOfContents(ByVal doc As Document, ByVal index As Integer)
			' Store the FieldStart nodes of TOC fields in the document for quick access.
			Dim fieldStarts As New ArrayList()
			' This is a list to store the nodes found inside the specified TOC. They will be removed
			' at thee end of this method.
			Dim nodeList As New ArrayList()

			For Each start As FieldStart In doc.GetChildNodes(NodeType.FieldStart, True)
				If start.FieldType = FieldType.FieldTOC Then
					' Add all FieldStarts which are of type FieldTOC.
					fieldStarts.Add(start)
				End If
			Next start

			' Ensure the TOC specified by the passed index exists.
			If index > fieldStarts.Count - 1 Then
				Throw New ArgumentOutOfRangeException("TOC index is out of range")
			End If

			Dim isRemoving As Boolean = True
			' Get the FieldStart of the specified TOC.
			Dim currentNode As Node = CType(fieldStarts(index), Node)

			Do While isRemoving
				' It is safer to store these nodes and delete them all at once later.
				nodeList.Add(currentNode)
				currentNode = currentNode.NextPreOrder(doc)

				' Once we encounter a FieldEnd node of type FieldTOC then we know we are at the end
				' of the current TOC and we can stop here.
				If currentNode.NodeType = NodeType.FieldEnd Then
					Dim fieldEnd As FieldEnd = CType(currentNode, FieldEnd)
					If fieldEnd.FieldType = FieldType.FieldTOC Then
						isRemoving = False
					End If
				End If
			Loop

			' Remove all nodes found in the specified TOC.
			For Each node As Node In nodeList
				node.Remove()
			Next node
		End Sub
		'ExEnd

		'ExStart
		'ExId:TCFieldsRangeReplace
		'ExSummary:Shows how to find and insert a TC field at text in a document. 
		<Test> _
		Public Sub InsertTCFieldsAtText()
			Dim doc As New Document()

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
				Dim builder As New DocumentBuilder(CType(args.MatchNode.Document, Document))
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
