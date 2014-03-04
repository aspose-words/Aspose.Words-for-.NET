' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Namespace ProcessComments
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The sample infrastructure.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Open the document.
			Dim doc As New Document(dataDir & "TestFile.doc")

			'ExStart
			'ExId:ProcessComments_Main
			'ExSummary: The demo-code that illustrates the methods for the comments extraction and removal.
			' Extract the information about the comments of all the authors.
			For Each comment As String In ExtractComments(doc)
				Console.Write(comment)
			Next comment

			' Remove comments by the "pm" author.
			RemoveComments(doc, "pm")
			Console.WriteLine("Comments from ""pm"" are removed!")

			' Extract the information about the comments of the "ks" author.
			For Each comment As String In ExtractComments(doc, "ks")
				Console.Write(comment)
			Next comment

			' Remove all comments.
			RemoveComments(doc)
			Console.WriteLine("All comments are removed!")

			' Save the document.
			doc.Save(dataDir & "Test File Out.doc")
			'ExEnd
		End Sub

		''' <param name="doc">The source document.</param>
		'ExStart
		'ExFor:Comment.Author
		'ExFor:Comment.DateTime
		'ExId:ProcessComments_Extract_All
		'ExSummary:Extracts the author name, date&time and text of all comments in the document.
		Private Shared Function ExtractComments(ByVal doc As Document) As ArrayList
			Dim collectedComments As New ArrayList()
			' Collect all comments in the document
			Dim comments As NodeCollection = doc.GetChildNodes(NodeType.Comment, True)
			' Look through all comments and gather information about them.
			For Each comment As Comment In comments
				collectedComments.Add(comment.Author & " " & comment.DateTime & " " & comment.ToString(SaveFormat.Text))
			Next comment
			Return collectedComments
		End Function
		'ExEnd

		''' <param name="doc">The source document.</param>
		''' <param name="authorName">The name of the comment's author.</param>
		'ExStart
		'ExId:ProcessComments_Extract_Author
		'ExSummary:Extracts the author name, date&time and text of the comments by the specified author.
		Private Shared Function ExtractComments(ByVal doc As Document, ByVal authorName As String) As ArrayList
			Dim collectedComments As New ArrayList()
			' Collect all comments in the document
			Dim comments As NodeCollection = doc.GetChildNodes(NodeType.Comment, True)
			' Look through all comments and gather information about those written by the authorName author.
			For Each comment As Comment In comments
				If comment.Author = authorName Then
					collectedComments.Add(comment.Author & " " & comment.DateTime & " " & comment.ToString(SaveFormat.Text))
				End If
			Next comment
			Return collectedComments
		End Function
		'ExEnd

		''' <param name="doc">The source document.</param>
		'ExStart
		'ExId:ProcessComments_Remove_All
		'ExSummary:Removes all comments in the document.
		Private Shared Sub RemoveComments(ByVal doc As Document)
			' Collect all comments in the document
			Dim comments As NodeCollection = doc.GetChildNodes(NodeType.Comment, True)
			' Remove all comments.
			comments.Clear()
		End Sub
		'ExEnd

		''' <param name="doc">The source document.</param>
		''' <param name="authorName">The name of the comment's author.</param>
		'ExStart
		'ExId:ProcessComments_Remove_Author
		'ExSummary:Removes comments by the specified author.
		Private Shared Sub RemoveComments(ByVal doc As Document, ByVal authorName As String)
			' Collect all comments in the document
			Dim comments As NodeCollection = doc.GetChildNodes(NodeType.Comment, True)
			' Look through all comments and remove those written by the authorName author.
			For i As Integer = comments.Count - 1 To 0 Step -1
				Dim comment As Comment = CType(comments(i), Comment)
				If comment.Author = authorName Then
					comment.Remove()
				End If
			Next i
		End Sub
		'ExEnd
	End Class
End Namespace