' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

'ExStart
'ExFor:NodeList
'ExFor:FieldStart
'ExId:ReplaceHyperlinks
'ExSummary:Finds all hyperlinks in a Word document and changes their URL and display name.


Imports Microsoft.VisualBasic
Imports System
Imports System.Text
Imports System.Text.RegularExpressions

Imports Aspose.Words
Imports Aspose.Words.Fields

Imports NUnit.Framework

'ExSkip

Namespace ApiExamples
	''' <summary>
	''' Shows how to replace hyperlinks in a Word document.
	''' </summary>
	<TestFixture> _
	Public Class ExReplaceHyperlinks
		Inherits ApiExampleBase
		''' <summary>
		''' Finds all hyperlinks in a Word document and changes their URL and display name.
		''' </summary>
		<Test> _
		Public Sub ReplaceHyperlinks()
			' Specify your document name here.
			Dim doc As New Document(MyDir & "ReplaceHyperlinks.doc")

			' Hyperlinks in a Word documents are fields, select all field start nodes so we can find the hyperlinks.
			Dim fieldStarts As NodeList = doc.SelectNodes("//FieldStart")
			For Each fieldStart As FieldStart In fieldStarts
				If fieldStart.FieldType.Equals(FieldType.FieldHyperlink) Then
					' The field is a hyperlink field, use the "facade" class to help to deal with the field.
					Dim hyperlink As New Hyperlink(fieldStart)

					' Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
					If hyperlink.IsLocal Then
						Continue For
					End If

					' The Hyperlink class allows to set the target URL and the display name 
					' of the link easily by setting the properties.
					hyperlink.Target = NewUrl
					hyperlink.Name = NewName
				End If
			Next fieldStart

			doc.Save(MyDir & "\Artifacts\ReplaceHyperlinks.doc")
		End Sub

		Private Const NewUrl As String = "http://www.aspose.com"
		Private Const NewName As String = "Aspose - The .NET & Java Component Publisher"
	End Class


	''' <summary>
	''' This "facade" class makes it easier to work with a hyperlink field in a Word document. 
	''' 
	''' A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words 
	''' consists of several nodes and it might be difficult to work with all those nodes directly. 
	''' Note this is a simple implementation and will work only if the hyperlink code and name 
	''' each consist of one Run only.
	''' 
	''' [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
	''' 
	''' The field code contains a string in one of these formats:
	''' HYPERLINK "url"
	''' HYPERLINK \l "bookmark name"
	''' 
	''' The field result contains text that is displayed to the user.
	''' </summary>
	Friend Class Hyperlink
		Friend Sub New(ByVal fieldStart As FieldStart)
			If fieldStart Is Nothing Then
				Throw New ArgumentNullException("fieldStart")
			End If
			If (Not fieldStart.FieldType.Equals(FieldType.FieldHyperlink)) Then
				Throw New ArgumentException("Field start type must be FieldHyperlink.")
			End If

			Me.mFieldStart = fieldStart

			' Find the field separator node.
			Me.mFieldSeparator = FindNextSibling(Me.mFieldStart, NodeType.FieldSeparator)
			If Me.mFieldSeparator Is Nothing Then
				Throw New InvalidOperationException("Cannot find field separator.")
			End If

			' Find the field end node. Normally field end will always be found, but in the example document 
			' there happens to be a paragraph break included in the hyperlink and this puts the field end 
			' in the next paragraph. It will be much more complicated to handle fields which span several 
			' paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
			Me.mFieldEnd = FindNextSibling(Me.mFieldSeparator, NodeType.FieldEnd)

			' Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs.
			Dim fieldCode As String = GetTextSameParent(Me.mFieldStart.NextSibling, Me.mFieldSeparator)
			Dim match As Match = gRegex.Match(fieldCode.Trim())
			Me.mIsLocal = (match.Groups(1).Length > 0) 'The link is local if \l is present in the field code.
			Me.mTarget = match.Groups(2).Value
		End Sub

		''' <summary>
		''' Gets or sets the display name of the hyperlink.
		''' </summary>
		Friend Property Name() As String
			Get
				Return GetTextSameParent(Me.mFieldSeparator, Me.mFieldEnd)
			End Get
			Set(ByVal value As String)
				' Hyperlink display name is stored in the field result which is a Run 
				' node between field separator and field end.
				Dim fieldResult As Run = CType(Me.mFieldSeparator.NextSibling, Run)
				fieldResult.Text = value

				' But sometimes the field result can consist of more than one run, delete these runs.
				RemoveSameParent(fieldResult.NextSibling, Me.mFieldEnd)
			End Set
		End Property

		''' <summary>
		''' Gets or sets the target url or bookmark name of the hyperlink.
		''' </summary>
		Friend Property Target() As String
			Get
				Dim dummy As String = Nothing ' This is needed to fool the C# to VB.NET converter.
				Return Me.mTarget
			End Get
			Set(ByVal value As String)
				Me.mTarget = value
				Me.UpdateFieldCode()
			End Set
		End Property

		''' <summary>
		''' True if the hyperlink's target is a bookmark inside the document. False if the hyperlink is a url.
		''' </summary>
		Friend Property IsLocal() As Boolean
			Get
				Return Me.mIsLocal
			End Get
			Set(ByVal value As Boolean)
				Me.mIsLocal = value
				Me.UpdateFieldCode()
			End Set
		End Property

		Private Sub UpdateFieldCode()
			' Field code is stored in a Run node between field start and field separator.
			Dim fieldCode As Run = CType(Me.mFieldStart.NextSibling, Run)
			fieldCode.Text = String.Format("HYPERLINK {0}""{1}""", (If((Me.mIsLocal), "\l ", "")), Me.mTarget)

			' But sometimes the field code can consist of more than one run, delete these runs.
			RemoveSameParent(fieldCode.NextSibling, Me.mFieldSeparator)
		End Sub

		''' <summary>
		''' Goes through siblings starting from the start node until it finds a node of the specified type or null.
		''' </summary>
		Private Shared Function FindNextSibling(ByVal startNode As Node, ByVal nodeType As NodeType) As Node
			Dim node As Node = startNode
			Do While node IsNot Nothing
				If node.NodeType.Equals(nodeType) Then
					Return node
				End If
				node = node.NextSibling
			Loop
			Return Nothing
		End Function

		''' <summary>
		''' Retrieves text from start up to but not including the end node.
		''' </summary>
		Private Shared Function GetTextSameParent(ByVal startNode As Node, ByVal endNode As Node) As String
			If (endNode IsNot Nothing) AndAlso (startNode.ParentNode IsNot endNode.ParentNode) Then
				Throw New ArgumentException("Start and end nodes are expected to have the same parent.")
			End If

			Dim builder As New StringBuilder()
			Dim child As Node = startNode
			Do While Not child.Equals(endNode)
				builder.Append(child.GetText())
				child = child.NextSibling
			Loop

			Return builder.ToString()
		End Function

		''' <summary>
		''' Removes nodes from start up to but not including the end node.
		''' Start and end are assumed to have the same parent.
		''' </summary>
		Private Shared Sub RemoveSameParent(ByVal startNode As Node, ByVal endNode As Node)
			If (endNode IsNot Nothing) AndAlso (startNode.ParentNode IsNot endNode.ParentNode) Then
				Throw New ArgumentException("Start and end nodes are expected to have the same parent.")
			End If

			Dim curChild As Node = startNode
			Do While (curChild IsNot Nothing) AndAlso (curChild IsNot endNode)
				Dim nextChild As Node = curChild.NextSibling
				curChild.Remove()
				curChild = nextChild
			Loop
		End Sub

		Private ReadOnly mFieldStart As Node
		Private ReadOnly mFieldSeparator As Node
		Private ReadOnly mFieldEnd As Node
		Private mIsLocal As Boolean
		Private mTarget As String

		''' <summary>
		''' RK I am notoriously bad at regexes. It seems I don't understand their way of thinking.
		''' </summary>
		Private Shared ReadOnly gRegex As New Regex("\S+" & "\s+" & "(?:""""\s+)?" & "(\\l\s+)?" & """" & 				 "([^""]+)" & 		 """"				)
	End Class
End Namespace
'ExEnd