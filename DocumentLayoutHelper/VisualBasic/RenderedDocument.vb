Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text

Imports Aspose.Words
Imports Aspose.Words.Layout

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
		''' <para>Whenever document is updated and new page layout model is created, a new enumerator must be used to access it.</para></remarks>
		Public Sub New(ByVal doc As Document)
			mEnumerator = New LayoutEnumerator(doc)
			ProcessLayoutElements(Me)
			CollectLinesAndAddToMarkers()
			LinkLinesToNodes(doc)
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
		''' Returns all lines of the specified document paragraph.
		''' </summary>
		''' <remarks>Note that this method sometimes won't return all lines if the paragraph is inside a table.</remarks>
		Public Function GetLinesOfParagraph(ByVal para As Paragraph) As LayoutCollection(Of RenderedLine)
			mEnumerator.MoveNode(para)
			Dim lines As List(Of RenderedLine) = New List(Of RenderedLine)()

			If mLayoutToLinesLookup.ContainsKey(mEnumerator.Current) Then
				lines = mLayoutToLinesLookup(mEnumerator.Current)
			End If

			Return New LayoutCollection(Of RenderedLine)(lines)
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
			CollectLinesOfMarkersCore(LayoutEntityType.HeaderFooter)
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
							End If
						Next span
					Next line
				Next story
			Next page
		End Sub

		Private Sub LinkLinesToNodes(ByVal doc As Document)
			For Each para As Paragraph In doc.GetChildNodes(NodeType.Paragraph, True)
				mEnumerator.MoveNode(para)

				If mLayoutToLinesLookup.ContainsKey(mEnumerator.Current) Then
					For Each entity As LayoutEntity In mLayoutToLinesLookup(mEnumerator.Current)
						CType(entity, NodeReferenceLayoutEntity).Paragraph = para
					Next entity
				End If
			Next para
		End Sub

		Private mEnumerator As LayoutEnumerator
		Private Shared mLayoutToLinesLookup As Dictionary(Of Object, List(Of RenderedLine)) = New Dictionary(Of Object, List(Of RenderedLine))()
	End Class
End Namespace
