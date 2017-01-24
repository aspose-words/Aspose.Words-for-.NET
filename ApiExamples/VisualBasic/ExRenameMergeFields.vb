' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

'ExStart
'ExId:RenameMergeFields
'ExSummary:Shows how to rename merge fields in a Word document.


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
	''' Shows how to rename merge fields in a Word document.
	''' </summary>
	<TestFixture> _
	Public Class ExRenameMergeFields
		Inherits ApiExampleBase
		''' <summary>
		''' Finds all merge fields in a Word document and changes their names.
		''' </summary>
		<Test> _
		Public Sub RenameMergeFields()
			' Specify your document name here.
			Dim doc As New Document(MyDir & "RenameMergeFields.doc")

			' Select all field start nodes so we can find the merge fields.
			Dim fieldStarts As NodeCollection = doc.GetChildNodes(NodeType.FieldStart, True)
			For Each fieldStart As FieldStart In fieldStarts
				If fieldStart.FieldType.Equals(FieldType.FieldMergeField) Then
					Dim mergeField As New MergeField(fieldStart)
					mergeField.Name = mergeField.Name & "_Renamed"
				End If
			Next fieldStart

			doc.Save(MyDir & "\Artifacts\RenameMergeFields.doc")
		End Sub
	End Class

	''' <summary>
	''' Represents a facade object for a merge field in a Microsoft Word document.
	''' </summary>
	Friend Class MergeField
		Friend Sub New(ByVal fieldStart As FieldStart)
			If fieldStart.Equals(Nothing) Then
				Throw New ArgumentNullException("fieldStart")
			End If
			If (Not fieldStart.FieldType.Equals(FieldType.FieldMergeField)) Then
				Throw New ArgumentException("Field start type must be FieldMergeField.")
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
		End Sub

		''' <summary>
		''' Gets or sets the name of the merge field.
		''' </summary>
		Friend Property Name() As String
			Get
				Return GetTextSameParent(Me.mFieldSeparator.NextSibling, Me.mFieldEnd).Trim("«"c, "»"c)
			End Get
			Set(ByVal value As String)
				' Merge field name is stored in the field result which is a Run 
				' node between field separator and field end.
				Dim fieldResult As Run = CType(Me.mFieldSeparator.NextSibling, Run)
				fieldResult.Text = String.Format("«{0}»", value)

				' But sometimes the field result can consist of more than one run, delete these runs.
				RemoveSameParent(fieldResult.NextSibling, Me.mFieldEnd)

				Me.UpdateFieldCode(value)
			End Set
		End Property

		Private Sub UpdateFieldCode(ByVal fieldName As String)
			' Field code is stored in a Run node between field start and field separator.
			Dim fieldCode As Run = CType(Me.mFieldStart.NextSibling, Run)
			Dim match As Match = gRegex.Match(fieldCode.Text)

			Dim newFieldCode As String = String.Format(" {0}{1} ", match.Groups("start").Value, fieldName)
			fieldCode.Text = newFieldCode

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

		Private Shared ReadOnly gRegex As New Regex("\s*(?<start>MERGEFIELD\s|)(\s|)(?<name>\S+)\s+")
	End Class
End Namespace
'ExEnd