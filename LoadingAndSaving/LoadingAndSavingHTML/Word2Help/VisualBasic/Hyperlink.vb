'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Text
Imports System.Text.RegularExpressions
Imports Aspose.Words
Imports Aspose.Words.Fields

Namespace Word2HelpExample
	''' <summary>
	''' This "facade" class makes it easier to work with a hyperlink field in a Word document. 
	''' 
	''' A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words 
	''' consists of several nodes and it might be difficult to work with all those nodes directly. 
	''' This is a simple implementation and will work only if the hyperlink code and name 
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
	Public Class Hyperlink
		Public Sub New(ByVal fieldStart As FieldStart)
			If fieldStart Is Nothing Then
				Throw New ArgumentNullException("fieldStart")
			End If
			If fieldStart.FieldType <> FieldType.FieldHyperlink Then
				Throw New ArgumentException("Field start type must be FieldHyperlink.")
			End If

			mFieldStart = fieldStart

			' Find field separator node.
			mFieldSeparator = FindNextSibling(mFieldStart, NodeType.FieldSeparator)
			If mFieldSeparator Is Nothing Then
				Throw New Exception("Cannot find field separator.")
			End If

			' Find field end node. Normally field end will always be found, but in the example document 
			' there happens to be a paragraph break included in the hyperlink and this puts the field end 
			' in the next paragraph. It will be much more complicated to handle fields which span several 
			' paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
			mFieldEnd = FindNextSibling(mFieldSeparator, NodeType.FieldEnd)

			' Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs.
			Dim fieldCode As String = GetTextSameParent(mFieldStart.NextSibling, mFieldSeparator)
			Dim match As Match = gRegex.Match(fieldCode.Trim())
			mIsLocal = (match.Groups(1).Length > 0) ' The link is local if \l is present in the field code.
			mTarget = match.Groups(2).Value
		End Sub

		''' <summary>
		''' Gets or sets the display name of the hyperlink.
		''' </summary>
		Public Property Name() As String
			Get
				Return GetTextSameParent(mFieldSeparator, mFieldEnd)
			End Get
			Set(ByVal value As String)
				' Hyperlink display name is stored in the field result which is a Run 
				' node between field separator and field end.
				Dim fieldResult As Run = CType(mFieldSeparator.NextSibling, Run)
				fieldResult.Text = value

				' But sometimes the field result can consist of more than one run, delete these runs.
				RemoveSameParent(fieldResult.NextSibling, mFieldEnd)
			End Set
		End Property

		''' <summary>
		''' Gets or sets the target url or bookmark name of the hyperlink.
		''' </summary>
		Public Property Target() As String
			Get
		Dim x As Integer = 0 ' RK This "fixes" the CSharp to VB.NET converter.
				Return mTarget
			End Get
			Set(ByVal value As String)
				mTarget = value
				UpdateFieldCode()
			End Set
		End Property

		''' <summary>
		''' True if the hyperlink's target is a bookmark inside the document. False if the hyperlink is a url.
		''' </summary>
		Public Property IsLocal() As Boolean
			Get
				Return mIsLocal
			End Get
			Set(ByVal value As Boolean)
				mIsLocal = value
				UpdateFieldCode()
			End Set
		End Property

		''' <summary>
		''' Updates the field code.
		''' </summary>
		Private Sub UpdateFieldCode()
			' Field code is stored in a Run node between field start and field separator.
			Dim fieldCode As Run = CType(mFieldStart.NextSibling, Run)
			fieldCode.Text = String.Format("HYPERLINK {0}""{1}""", (If((mIsLocal), "\l ", "")), mTarget)

			' But sometimes the field code can consist of more than one run, delete these runs.
			RemoveSameParent(fieldCode.NextSibling, mFieldSeparator)
		End Sub

		''' <summary>
		''' Goes through siblings starting from the start node until it finds a node of the specified type or null.
		''' </summary>
		Private Shared Function FindNextSibling(ByVal start As Node, ByVal nodeType As NodeType) As Node
			Dim node As Node = start
			Do While node IsNot Nothing
				If node.NodeType = nodeType Then
					Return node
				End If
				node = node.NextSibling
			Loop
			Return Nothing
		End Function

		''' <summary>
		''' Retrieves text from start up to but not including the end node.
		''' </summary>
		Private Shared Function GetTextSameParent(ByVal start As Node, ByVal [end] As Node) As String
			If ([end] IsNot Nothing) AndAlso (start.ParentNode IsNot [end].ParentNode) Then
				Throw New ArgumentException("Start and end nodes are expected to have the same parent.")
			End If

			Dim builder As New StringBuilder()
			Dim child As Node = start
			Do While child IsNot [end]
				builder.Append(child.GetText())
				child = child.NextSibling
			Loop
			Return builder.ToString()
		End Function

		''' <summary>
		''' Removes nodes from start up to but not including the end node.
		''' Start and end are assumed to have the same parent.
		''' </summary>
		Private Shared Sub RemoveSameParent(ByVal start As Node, ByVal [end] As Node)
			If ([end] IsNot Nothing) AndAlso (start.ParentNode IsNot [end].ParentNode) Then
				Throw New ArgumentException("Start and end nodes are expected to have the same parent.")
			End If

			Dim curChild As Node = start
			Do While curChild IsNot [end]
				Dim nextChild As Node = curChild.NextSibling
				curChild.Remove()
				curChild = nextChild
			Loop
		End Sub

		Private ReadOnly mFieldStart As Node
		Private ReadOnly mFieldSeparator As Node
		Private ReadOnly mFieldEnd As Node
		Private mTarget As String
		Private mIsLocal As Boolean

		Private Shared ReadOnly gRegex As New Regex("\S+" & "\s+" & "(?:""""\s+)?" & "(\\l\s+)?" & """" & "([^""]+)" & """")
	End Class
End Namespace