' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Text.RegularExpressions

Imports Aspose.Words
Imports Aspose.Words.MailMerging

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExInsertDocument
		Inherits ApiExampleBase
		'ExStart
		'ExFor:Paragraph.IsEndOfSection
		'ExId:InsertDocumentMain
		'ExSummary:This is a method that inserts contents of one document at a specified location in another document.
		''' <summary>
		''' Inserts content of the external document after the specified node.
		''' Section breaks and section formatting of the inserted document are ignored.
		''' </summary>
		''' <param name="insertAfterNode">Node in the destination document after which the content 
		''' should be inserted. This node should be a block level node (paragraph or table).</param>
		''' <param name="srcDoc">The document to insert.</param>
		Private Shared Sub InsertDocument(ByVal insertAfterNode As Node, ByVal srcDoc As Document)
			' Make sure that the node is either a paragraph or table.
			If ((Not insertAfterNode.NodeType.Equals(NodeType.Paragraph))) And ((Not insertAfterNode.NodeType.Equals(NodeType.Table))) Then
				Throw New ArgumentException("The destination node should be either a paragraph or table.")
			End If

			' We will be inserting into the parent of the destination paragraph.
			Dim dstStory As CompositeNode = insertAfterNode.ParentNode

			' This object will be translating styles and lists during the import.
			Dim importer As New NodeImporter(srcDoc, insertAfterNode.Document, ImportFormatMode.KeepSourceFormatting)

			' Loop through all sections in the source document.
			For Each srcSection As Section In srcDoc.Sections
				' Loop through all block level nodes (paragraphs and tables) in the body of the section.
				For Each srcNode As Node In srcSection.Body
					' Let's skip the node if it is a last empty paragraph in a section.
					If srcNode.NodeType.Equals(NodeType.Paragraph) Then
						Dim para As Paragraph = CType(srcNode, Paragraph)
						If para.IsEndOfSection AndAlso (Not para.HasChildNodes) Then
							Continue For
						End If
					End If

					' This creates a clone of the node, suitable for insertion into the destination document.
					Dim newNode As Node = importer.ImportNode(srcNode, True)

					' Insert new node after the reference node.
					dstStory.InsertAfter(newNode, insertAfterNode)
					insertAfterNode = newNode
				Next srcNode
			Next srcSection
		End Sub
		'ExEnd

		<Test> _
		Public Sub InsertDocumentAtBookmark()
			'ExStart
			'ExId:InsertDocumentAtBookmark
			'ExSummary:Invokes the InsertDocument method shown above to insert a document at a bookmark.
			Dim mainDoc As New Document(MyDir & "InsertDocument1.doc")
			Dim subDoc As New Document(MyDir & "InsertDocument2.doc")

			Dim bookmark As Bookmark = mainDoc.Range.Bookmarks("insertionPlace")
			InsertDocument(bookmark.BookmarkStart.ParentNode, subDoc)

			mainDoc.Save(MyDir & "\Artifacts\InsertDocumentAtBookmark.doc")
			'ExEnd
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub InsertDocumentAtMailMergeCaller()
			Me.InsertDocumentAtMailMerge()
		End Sub

		'ExStart
		'ExFor:CompositeNode.HasChildNodes
		'ExId:InsertDocumentAtMailMerge
		'ExSummary:Demonstrates how to use the InsertDocument method to insert a document into a merge field during mail merge.
		Public Sub InsertDocumentAtMailMerge()
			' Open the main document.
			Dim mainDoc As New Document(MyDir & "InsertDocument1.doc")

			' Add a handler to MergeField event
			mainDoc.MailMerge.FieldMergingCallback = New InsertDocumentAtMailMergeHandler()

			' The main document has a merge field in it called "Document_1".
			' The corresponding data for this field contains fully qualified path to the document
			' that should be inserted to this field.
			mainDoc.MailMerge.Execute(New String() { "Document_1" }, New String() { MyDir & "InsertDocument2.doc" })

			mainDoc.Save(MyDir & "\Artifacts\InsertDocumentAtMailMerge.doc")
		End Sub

		Private Class InsertDocumentAtMailMergeHandler
			Implements IFieldMergingCallback
			''' <summary>
			''' This handler makes special processing for the "Document_1" field.
			''' The field value contains the path to load the document. 
			''' We load the document and insert it into the current merge field.
			''' </summary>
			Private Sub IFieldMergingCallback_FieldMerging(ByVal e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
				If e.DocumentFieldName = "Document_1" Then
					' Use document builder to navigate to the merge field with the specified name.
					Dim builder As New DocumentBuilder(e.Document)
					builder.MoveToMergeField(e.DocumentFieldName)

					' The name of the document to load and insert is stored in the field value.
					Dim subDoc As New Document(CStr(e.FieldValue))

					' Insert the document.
					InsertDocument(builder.CurrentParagraph, subDoc)

					' The paragraph that contained the merge field might be empty now and you probably want to delete it.
					If (Not builder.CurrentParagraph.HasChildNodes) Then
						builder.CurrentParagraph.Remove()
					End If

					' Indicate to the mail merge engine that we have inserted what we wanted.
					e.Text = Nothing
				End If
			End Sub

			Private Sub ImageFieldMerging(ByVal args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
				' Do nothing.
			End Sub
		End Class
		'ExEnd

		'ExStart
		'ExId:InsertDocumentAtMailMergeBlob
		'ExSummary:A slight variation to the above example to load a document from a BLOB database field instead of a file.
		Private Class InsertDocumentAtMailMergeBlobHandler
			Implements IFieldMergingCallback
			''' <summary>
			''' This handler makes special processing for the "Document_1" field.
			''' The field value contains the path to load the document. 
			''' We load the document and insert it into the current merge field.
			''' </summary>
			Private Sub IFieldMergingCallback_FieldMerging(ByVal e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
				If e.DocumentFieldName = "Document_1" Then
					' Use document builder to navigate to the merge field with the specified name.
					Dim builder As New DocumentBuilder(e.Document)
					builder.MoveToMergeField(e.DocumentFieldName)

					' Load the document from the blob field.
					Dim stream As New MemoryStream(CType(e.FieldValue, Byte()))
					Dim subDoc As New Document(stream)

					' Insert the document.
					InsertDocument(builder.CurrentParagraph, subDoc)

					' The paragraph that contained the merge field might be empty now and you probably want to delete it.
					If (Not builder.CurrentParagraph.HasChildNodes) Then
						builder.CurrentParagraph.Remove()
					End If

					' Indicate to the mail merge engine that we have inserted what we wanted.
					e.Text = Nothing
				End If
			End Sub

			Private Sub ImageFieldMerging(ByVal args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
				' Do nothing.
			End Sub
		End Class
		'ExEnd

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub InsertDocumentAtReplaceCaller()
			Me.InsertDocumentAtReplace()
		End Sub

		'ExStart
		'ExFor:Range.Replace(Regex,IReplacingCallback,Boolean)
		'ExFor:IReplacingCallback
		'ExFor:ReplaceAction
		'ExFor:IReplacingCallback.Replacing
		'ExFor:ReplacingArgs
		'ExFor:ReplacingArgs.MatchNode
		'ExId:InsertDocumentAtReplace
		'ExSummary:Shows how to insert content of one document into another during a customized find and replace operation.
		Public Sub InsertDocumentAtReplace()
			Dim mainDoc As New Document(MyDir & "InsertDocument1.doc")
			mainDoc.Range.Replace(New Regex("\[MY_DOCUMENT\]"), New InsertDocumentAtReplaceHandler(), False)
			mainDoc.Save(MyDir & "\Artifacts\InsertDocumentAtReplace.doc")
		End Sub

		Private Class InsertDocumentAtReplaceHandler
			Implements IReplacingCallback
			Private Function IReplacingCallback_Replacing(ByVal e As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
				Dim subDoc As New Document(MyDir & "InsertDocument2.doc")

				' Insert a document after the paragraph, containing the match text.
				Dim para As Paragraph = CType(e.MatchNode.ParentNode, Paragraph)
				InsertDocument(para, subDoc)

				' Remove the paragraph with the match text.
				para.Remove()

				Return ReplaceAction.Skip
			End Function
		End Class
		'ExEnd
	End Class
End Namespace
