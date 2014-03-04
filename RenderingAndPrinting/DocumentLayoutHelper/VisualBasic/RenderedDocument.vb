' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text

Imports Aspose.Words
Imports Aspose.Words.Layout
Imports Aspose.Words.Tables

Namespace Aspose.Words.Layout
	''' <summary>
	''' Provides an API wrapper for the LayoutEnumerator class to access the page layout entities of a document presented in
	''' a object model like design.
	''' </summary>
	Public Class RenderedDocument
		Inherits LayoutEntity
		''' <summary>
		''' Creates a new instance from the supplied Aspose.Words.Document class.
		''' </summary>
		''' <param name="document">A document whose page layout model to enumerate.</param>
		''' <remarks><para>If page layout model of the document hasn't been built the enumerator calls <see cref="Document.UpdatePageLayout"/> to build it.</para>
		''' <para>Whenever document is updated and new page layout model is created, a new RenderedDocument instance must be used to access the changes.</para></remarks>
		Public Sub New(ByVal doc As Document)
			mLayoutCollector = New LayoutCollector(doc)
			mEnumerator = New LayoutEnumerator(doc)
			ProcessLayoutElements(Me)
			CollectLinesAndAddToMarkers()
			LinkLayoutMarkersToNodes(doc)
		End Sub

		''' <summary>
		''' Provides access to the pages of a document.
		''' </summary>
		Public ReadOnly Property Pages() As LayoutCollection(Of RenderedPage)
			Get
				Return GetChildNodes(Of RenderedPage)()
			End Get
		End Property

		''' <summary>
		''' Returns all the layout entities of the specified node.
		''' </summary>
		''' <remarks>Note that this method does not work with Run nodes or nodes in the header or footer.</remarks>
		Public Function GetLayoutEntitiesOfNode(ByVal node As Node) As LayoutCollection(Of LayoutEntity)
			If (Not mLayoutCollector.Document.Equals(node.Document)) Then
				Throw New ArgumentException("Node does not belong to the same document which was rendered.")
			End If

			If node.NodeType = NodeType.Document Then
				Return New LayoutCollection(Of LayoutEntity)(mChildEntities)
			End If

			Dim entities As List(Of LayoutEntity) = New List(Of LayoutEntity)()

			' Retrieve all entities from the layout document (inversion of LayoutEntityType.None).
			For Each entity As LayoutEntity In GetChildEntities((Not LayoutEntityType.None), True)
				If entity.ParentNode Is node Then
					entities.Add(entity)
				End If

				' There is no table entity in rendered output so manually check if rows belong to a table node.
				If entity.Type = LayoutEntityType.Row Then
					Dim row As RenderedRow = CType(entity, RenderedRow)
					If row.Table Is node Then
						entities.Add(entity)
					End If
				End If
			Next entity

			Return New LayoutCollection(Of LayoutEntity)(entities)
		End Function

		Private Sub ProcessLayoutElements(ByVal current As LayoutEntity)
			Do
				Dim child As LayoutEntity = current.AddChildEntity(mEnumerator)

				If mEnumerator.MoveFirstChild() Then
					current = child

					ProcessLayoutElements(current)
					mEnumerator.MoveParent()

					current = current.Parent
				End If
			Loop While mEnumerator.MoveNext()
		End Sub

		Private Sub CollectLinesAndAddToMarkers()
			CollectLinesOfMarkersCore(LayoutEntityType.Column)
			CollectLinesOfMarkersCore(LayoutEntityType.Comment)
		End Sub

		Private Sub CollectLinesOfMarkersCore(ByVal type As LayoutEntityType)
			Dim collectedLines As List(Of RenderedLine) = New List(Of RenderedLine)()

			For Each page As RenderedPage In Pages
				For Each story As LayoutEntity In page.GetChildEntities(type, False)
					For Each line As RenderedLine In story.GetChildEntities(LayoutEntityType.Line, True)
						collectedLines.Add(line)
						For Each span As RenderedSpan In line.Spans
							If span.Kind = "PARAGRAPH" OrElse span.Kind = "ROW" OrElse span.Kind = "CELL" OrElse span.Kind = "SECTION" Then
								mLayoutToLinesLookup.Add(span.LayoutObject, collectedLines)
								collectedLines = New List(Of RenderedLine)()
							Else
								mLayoutToSpanLookup.Add(span.LayoutObject, span)
							End If
						Next span
					Next line
				Next story
			Next page
		End Sub

		Private Sub LinkLayoutMarkersToNodes(ByVal doc As Document)
			For Each node As Node In doc.GetChildNodes(NodeType.Any, True)
				Select Case node.NodeType
					Case NodeType.Paragraph
						For Each line As RenderedLine In GetLinesOfNode(node)
							line.SetParentNode(node)
						Next line

					Case NodeType.Row
						For Each line As RenderedLine In GetLinesOfNode(node)
							line.SetParentNode((CType(node, Row)).LastCell.LastParagraph)
						Next line

					Case Else
						If mLayoutCollector.GetEntity(node) IsNot Nothing Then
							mLayoutToSpanLookup(mLayoutCollector.GetEntity(node)).SetParentNode(node)
						End If
				End Select
			Next node
		End Sub

		Private Function GetLinesOfNode(ByVal node As Node) As List(Of RenderedLine)
			Dim lines As List(Of RenderedLine) = New List(Of RenderedLine)()
			Dim nodeEntity As Object = mLayoutCollector.GetEntity(node)

			If nodeEntity IsNot Nothing AndAlso mLayoutToLinesLookup.ContainsKey(nodeEntity) Then
				lines = mLayoutToLinesLookup(nodeEntity)
			End If

			Return lines
		End Function

		Private mLayoutCollector As LayoutCollector
		Private mEnumerator As LayoutEnumerator
		Private Shared mLayoutToLinesLookup As Dictionary(Of Object, List(Of RenderedLine)) = New Dictionary(Of Object, List(Of RenderedLine))()
		Private Shared mLayoutToSpanLookup As Dictionary(Of Object, RenderedSpan) = New Dictionary(Of Object, RenderedSpan)()
	End Class
End Namespace
