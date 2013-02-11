Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Drawing

Imports Aspose.Words.Tables
Imports Aspose.Words.Layout

Namespace Aspose.Words.Layout
	''' <summary>
	''' Provides the base class for rendered elements of a document.
	''' </summary>
	Public MustInherit Class LayoutEntity
		Protected Sub New()
		End Sub

		''' <summary>
		''' Gets the 1-based index of a page which contains the rendered entity.
		''' </summary>
		Public ReadOnly Property PageIndex() As Integer
			Get
				Return mPageIndex
			End Get
		End Property

		''' <summary>
		''' Returns bounding rectangle of the entity relative to the page top left corner (in points).
		''' </summary>
		Public ReadOnly Property Rectangle() As RectangleF
			Get
				Return mRectangle
			End Get
		End Property

		''' <summary>
		''' Gets the type of this layout entity.
		''' </summary>
		Public ReadOnly Property Type() As LayoutEntityType
			Get
				Return mType
			End Get
		End Property

		''' <summary>
		''' Exports the contents of the entity into a string in plain text format.
		''' </summary>
		Public Overridable ReadOnly Property Text() As String
			Get
				Dim builder As New StringBuilder()
				For Each entity As LayoutEntity In mChildEntities
					builder.Append(entity.Text)
				Next entity

				Return builder.ToString()
			End Get
		End Property

		''' <summary>
		''' Gets the immediate parent of this entity.
		''' </summary>
		Public ReadOnly Property Parent() As LayoutEntity
			Get
				Return mParent
			End Get
		End Property

		''' <summary>
		''' Reserved for internal use.
		''' </summary>
		Private privateLayoutObject As Object
		Friend Property LayoutObject() As Object
			Get
				Return privateLayoutObject
			End Get
			Set(ByVal value As Object)
				privateLayoutObject = value
			End Set
		End Property

		''' <summary>
		''' Reserved for internal use.
		''' </summary>
		Friend Function AddChildEntity(ByVal it As LayoutEnumerator) As LayoutEntity
			Dim child As LayoutEntity = CreateLayoutEntityFromType(it)
			mChildEntities.Add(child)

			Return child
		End Function

		Private Function CreateLayoutEntityFromType(ByVal it As LayoutEnumerator) As LayoutEntity
			Dim childEntity As LayoutEntity
			Select Case it.Type
				Case LayoutEntityType.Cell
					childEntity = New RenderedCell()
				Case LayoutEntityType.Column
					childEntity = New RenderedColumn()
				Case LayoutEntityType.Comment
					childEntity = New RenderedComment()
				Case LayoutEntityType.Endnote
					childEntity = New RenderedEndnote()
				Case LayoutEntityType.Footnote
					childEntity = New RenderedFootnote()
				Case LayoutEntityType.HeaderFooter
					childEntity = New RenderedHeaderFooter()
				Case LayoutEntityType.Line
					childEntity = New RenderedLine()
				Case LayoutEntityType.NoteSeparator
					childEntity = New RenderedNoteSeparator()
				Case LayoutEntityType.Page
					childEntity = New RenderedPage()
				Case LayoutEntityType.Row
					childEntity = New RenderedRow()
				Case LayoutEntityType.Span
					childEntity = New RenderedSpan(it.Text)
				Case LayoutEntityType.TextBox
					childEntity = New RenderedTextBox()
				Case Else
					Throw New InvalidOperationException("Unknown layout type")
			End Select

			childEntity.mKind = it.Kind
			childEntity.mPageIndex = it.PageIndex
			childEntity.mRectangle = it.Rectangle
			childEntity.mType = it.Type
			childEntity.LayoutObject = it.Current
			childEntity.mParent = Me

			Return childEntity
		End Function

		''' <summary>
		''' Returns a collection of child entities which match the specified type.
		''' </summary>
		''' <param name="type">Specifies the type of entities to select.</param>
		''' <param name="isDeep">True to select from all child entities recursively. False to select only among immediate children</param>
		Public Function GetChildEntities(ByVal type As LayoutEntityType, ByVal isDeep As Boolean) As LayoutCollection(Of LayoutEntity)
			Dim childList As List(Of LayoutEntity) = New List(Of LayoutEntity)()

			For Each entity As LayoutEntity In mChildEntities
				If entity.Type = type Then
					childList.Add(entity)
				End If

				If isDeep Then
					childList.AddRange(CType(entity.GetChildEntities(type, True), IEnumerable(Of LayoutEntity)))
				End If
			Next entity

			Return New LayoutCollection(Of LayoutEntity)(childList)
		End Function

		Protected Function GetChildNodes(Of T As {LayoutEntity, New})() As LayoutCollection(Of T)
			Dim obj As New T()
			Dim childList As List(Of T) = New List(Of T)()

			For Each entity As LayoutEntity In mChildEntities
				If entity.GetType() Is obj.GetType() Then
					childList.Add(CType(entity, T))
				End If
			Next entity

			Return New LayoutCollection(Of T)(childList)
		End Function

		Protected mKind As String
		Protected mPageIndex As Integer
		Protected mRectangle As RectangleF
		Protected mType As LayoutEntityType
		Protected mParent As LayoutEntity
		Protected mChildEntities As List(Of LayoutEntity) = New List(Of LayoutEntity)()
	End Class

	''' <summary>
	''' Represents a generic collection of layout entity types.
	''' </summary>
	Public Class LayoutCollection(Of T As LayoutEntity)
		Implements IEnumerable(Of T)
		''' <summary>
		''' Reserved for internal use.
		''' </summary>
		Friend Sub New(ByVal baseList As List(Of T))
			mBaseList = baseList
		End Sub

		''' <summary>
		''' Provides a simple "foreach" style iteration over the collection of nodes. 
		''' </summary>
		Private Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
			Return mBaseList.GetEnumerator()
		End Function

		''' <summary>
		''' Provides a simple "foreach" style iteration over the collection of nodes. 
		''' </summary>
		Private Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
			Return mBaseList.GetEnumerator()
		End Function

		''' <summary>
		''' Returns the first entity in the collection.
		''' </summary>
		Public ReadOnly Property First() As T
			Get
				If mBaseList.Count > 0 Then
					Return mBaseList(0)
				Else
					Return Nothing
				End If
			End Get
		End Property

		''' <summary>
		''' Returns the last entity in the collection.
		''' </summary>
		Public ReadOnly Property Last() As T
			Get
				If mBaseList.Count > 0 Then
					Return mBaseList(mBaseList.Count - 1)
				Else
					Return Nothing
				End If
			End Get
		End Property

		''' <summary>
		''' Retrieves the entity at the given index. 
		''' </summary>
		''' <remarks><para>The index is zero-based.</para>
		''' <para>If index is greater than or equal to the number of items in the list, this returns a null reference.</para></remarks>
		Default Public ReadOnly Property Item(ByVal index As Integer) As T
			Get
				Return mBaseList(index)
			End Get
		End Property

		''' <summary>
		''' Gets the number of entities in the collection.
		''' </summary>
		Public ReadOnly Property Count() As Integer
			Get
				Return mBaseList.Count
			End Get
		End Property

		Private mBaseList As List(Of T)
	End Class

	''' <summary>
	''' Represents an entity that contains lines and rows.
	''' </summary>
	Public MustInherit Class StoryLayoutEntity
		Inherits LayoutEntity
		''' <summary>
		''' Provides access to the lines of a story.
		''' </summary>
		Public ReadOnly Property Lines() As LayoutCollection(Of RenderedLine)
			Get
				Return GetChildNodes(Of RenderedLine)()
			End Get
		End Property

		''' <summary>
		''' Provides access to the row entities of a table.
		''' </summary>
		Public ReadOnly Property Rows() As LayoutCollection(Of RenderedRow)
			Get
				Return GetChildNodes(Of RenderedRow)()
			End Get
		End Property
	End Class

	''' <summary>
	''' Represents an entity that has a reference to a document node.
	''' </summary>
	Public MustInherit Class NodeReferenceLayoutEntity
		Inherits LayoutEntity
		''' <summary>
		''' Returns the document paragraph that corresponds to the layout entity.  
		''' </summary>
		''' <remarks>This property may return null for some lines such as those inside the header or footer.</remarks>
		Private privateParagraph As Paragraph
		Public Property Paragraph() As Paragraph
			Get
				Return privateParagraph
			End Get
			Set(ByVal value As Paragraph)
				privateParagraph = value
			End Set
		End Property
	End Class

	''' <summary>
	''' Represents line of characters of text and inline objects.
	''' </summary>
	Public Class RenderedLine
		Inherits NodeReferenceLayoutEntity
		Public Overrides ReadOnly Property Text() As String
			Get
				Return MyBase.Text & Environment.NewLine
			End Get
		End Property

		''' <summary>
		''' Provides access to the spans of the line.
		''' </summary>
		Public ReadOnly Property Spans() As LayoutCollection(Of RenderedSpan)
			Get
				Return GetChildNodes(Of RenderedSpan)()
			End Get
		End Property
	End Class

	''' <summary>
	''' Represents one or more characters in a line.
	''' This include special characters like field start/end markers, bookmarks and comments.
	''' </summary>
	Public Class RenderedSpan
		Inherits LayoutEntity
		Public Sub New()
		End Sub

		Friend Sub New(ByVal text As String)
			mText = text
		End Sub

		''' <summary>
		''' Gets kind of the span. This cannot be null.
		''' </summary>
		''' <remarks>This is a more specific type of the current entity, e.g. bookmark span has Span type and
		''' may have either a BOOKMARKSTART or BOOKMARKEND kind.</remarks>
		Public ReadOnly Property Kind() As String
			Get
				Return mKind
			End Get
		End Property

		''' <summary>
		''' Exports the contents of the entity into a string in plain text format.
		''' </summary>
		Public Overrides ReadOnly Property Text() As String
			Get
				Return mText
			End Get
		End Property

		Private mText As String
	End Class

	''' <summary>
	''' Represents the header/footer content on a page.
	''' </summary>
	Public Class RenderedHeaderFooter
		Inherits StoryLayoutEntity
		''' <summary>
		''' Returns the type of the header or footer.
		''' </summary>
		Public ReadOnly Property Kind() As String
			Get
				Return mKind
			End Get
		End Property
	End Class

	''' <summary>
	''' Represents page of a document.
	''' </summary>
	Public Class RenderedPage
		Inherits LayoutEntity
		''' <summary>
		''' Provides access to the columns of the page.
		''' </summary>
		Public ReadOnly Property Columns() As LayoutCollection(Of RenderedColumn)
			Get
				Return GetChildNodes(Of RenderedColumn)()
			End Get
		End Property

		''' <summary>
		''' Provides access to the header and footers of the page.
		''' </summary>
		Public ReadOnly Property HeaderFooters() As LayoutCollection(Of RenderedHeaderFooter)
			Get
				Return GetChildNodes(Of RenderedHeaderFooter)()
			End Get
		End Property

		''' <summary>
		''' Provides access to the comments of the page.
		''' </summary>
		Public ReadOnly Property Comments() As LayoutCollection(Of RenderedComment)
			Get
				Return GetChildNodes(Of RenderedComment)()
			End Get
		End Property
	End Class

	''' <summary>
	''' Represents a table row.
	''' </summary>
	Public Class RenderedRow
		Inherits LayoutEntity
		''' <summary>
		''' Provides access to the cells of the row.
		''' </summary>
		Public ReadOnly Property Cells() As LayoutCollection(Of RenderedCell)
			Get
				Return GetChildNodes(Of RenderedCell)()
			End Get
		End Property
	End Class

	''' <summary>
	''' Represents a column of text on a page.
	''' </summary>
	Public Class RenderedColumn
		Inherits StoryLayoutEntity
		''' <summary>
		''' Provides access to the footnotes of the page.
		''' </summary>
		Public ReadOnly Property Footnotes() As LayoutCollection(Of RenderedFootnote)
			Get
				Return GetChildNodes(Of RenderedFootnote)()
			End Get
		End Property

		''' <summary>
		''' Provides access to the endnotes of the page.
		''' </summary>
		Public ReadOnly Property Endnotes() As LayoutCollection(Of RenderedEndnote)
			Get
				Return GetChildNodes(Of RenderedEndnote)()
			End Get
		End Property

		''' <summary>
		''' Provides access to the note separators of the page.
		''' </summary>
		Public ReadOnly Property NoteSeparators() As LayoutCollection(Of RenderedNoteSeparator)
			Get
				Return GetChildNodes(Of RenderedNoteSeparator)()
			End Get
		End Property
	End Class

	''' <summary>
	''' Represents a table cell.
	''' </summary>
	Public Class RenderedCell
		Inherits StoryLayoutEntity
	End Class

	''' <summary>
	''' Represents placeholder for footnote content.
	''' </summary>
	Public Class RenderedFootnote
		Inherits StoryLayoutEntity
	End Class

	''' <summary>
	''' Represents placeholder for endnote content.
	''' </summary>
	Public Class RenderedEndnote
		Inherits StoryLayoutEntity
	End Class

	''' <summary>
	''' Represents text area inside of a shape.
	''' </summary>
	Public Class RenderedTextBox
		Inherits StoryLayoutEntity
	End Class

	''' <summary>
	''' Represents placeholder for comment content.
	''' </summary>
	Public Class RenderedComment
		Inherits StoryLayoutEntity
	End Class

	''' <summary>
	''' Represents footnote/endnote separator.
	''' </summary>
	Public Class RenderedNoteSeparator
		Inherits StoryLayoutEntity
	End Class
End Namespace
