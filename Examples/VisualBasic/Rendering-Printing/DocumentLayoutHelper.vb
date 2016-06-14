Imports Microsoft.VisualBasic
Imports System
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Layout
Imports Aspose.Words.Tables
Imports System.Drawing
Imports System.Text

Public Class DocumentLayoutHelper
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        Dim doc As New Document(dataDir & "TestFile.docx")

        ' This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for 
        ' the LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.

        ' Create a new RenderedDocument class from a Document object.
        Dim layoutDoc As New RenderedDocument(doc)

        ' The following examples demonstrate how to use the wrapper API. 
        ' This snippet returns the third line of the first page and prints the line of text to the console.
        Dim line As RenderedLine = layoutDoc.Pages(0).Columns(0).Lines(2)
        Console.WriteLine("Line: " & line.Text)

        ' With a rendered line the original paragraph in the document object model can be returned.
        Dim para As Paragraph = line.Paragraph
        Console.WriteLine("Paragraph text: " & para.Range.Text)

        ' Retrieve all the text that appears of the first page in plain text format (including headers and footers).
        Dim pageText As String = layoutDoc.Pages(0).Text
        Console.WriteLine()

        ' Loop through each page in the document and print how many lines appear on each page.
        For Each page As RenderedPage In layoutDoc.Pages
            Dim lines As LayoutCollection(Of LayoutEntity) = page.GetChildEntities(LayoutEntityType.Line, True)
            Console.WriteLine("Page {0} has {1} lines.", page.PageIndex, lines.Count)
        Next page

        ' This method provides a reverse lookup of layout entities for any given node (with the exception of runs and nodes in the
        ' header and footer).
        Console.WriteLine()
        Console.WriteLine("The lines of the second paragraph:")
        For Each paragraphLine As RenderedLine In layoutDoc.GetLayoutEntitiesOfNode(doc.FirstSection.Body.Paragraphs(1))
            Console.WriteLine(String.Format("""{0}""", paragraphLine.Text.Trim()))
            Console.WriteLine(paragraphLine.Rectangle.ToString())
            Console.WriteLine()
        Next paragraphLine

        Console.WriteLine(vbNewLine + "Document layout helper example ran successfully.")
    End Sub
End Class

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
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    ''' <remarks>This property may return null for spans that originate from Run nodes or nodes that are inside the header or footer.</remarks>
    Public Overridable ReadOnly Property ParentNode() As Node
        Get
            Return mParentNode
        End Get
    End Property

    ''' <summary>
    ''' Internal method separate from ParentNode property to make code autoportable to VB.NET.
    ''' </summary>
    Friend Overridable Sub SetParentNode(ByVal value As Node)
        mParentNode = value
    End Sub

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
            If (entity.Type And type) = entity.Type Then
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
    Protected mParentNode As Node
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
            If index < mBaseList.Count Then
                Return mBaseList(index)
            Else
                Return Nothing
            End If
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
''' Represents line of characters of text and inline objects.
''' </summary>
Public Class RenderedLine
    Inherits LayoutEntity
    ''' <summary>
    ''' Exports the contents of the entity into a string in plain text format.
    ''' </summary>
    Public Overrides ReadOnly Property Text() As String
        Get
            Return MyBase.Text & Environment.NewLine
        End Get
    End Property

    ''' <summary>
    ''' Returns the paragraph that corresponds to the layout entity.  
    ''' </summary>
    ''' <remarks>This property may return null for some lines such as those inside the header or footer.</remarks>
    Public ReadOnly Property Paragraph() As Paragraph
        Get
            Return CType(ParentNode, Paragraph)
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
''' This include special characters like field start/end markers, bookmarks, shapes and comments.
''' </summary>
Public Class RenderedSpan
    Inherits LayoutEntity
    Public Sub New()
    End Sub

    Friend Sub New(ByVal text As String)
        ' Assign empty text if the span text is null (this can happen with shape spans).
        mText = If(text IsNot Nothing, text, String.Empty)
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

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    ''' <remarks>This property returns null for spans that originate from Run nodes or nodes that are inside the header or footer.</remarks>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Return mParentNode
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

    ''' <summary>
    ''' Returns the section that corresponds to the layout entity.  
    ''' </summary>
    Public ReadOnly Property Section() As Section
        Get
            Return CType(ParentNode, Section)
        End Get
    End Property

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Return Columns.First.GetChildEntities(LayoutEntityType.Line, True).First.ParentNode.GetAncestor(NodeType.Section)
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

    ''' <summary>
    ''' Returns the row that corresponds to the layout entity.  
    ''' </summary>
    ''' <remarks>This property may return null for some rows such as those inside the header or footer.</remarks>
    Public ReadOnly Property Row() As Row
        Get
            Return CType(ParentNode, Row)
        End Get
    End Property

    ''' <summary>
    ''' Returns the table that corresponds to the layout entity.  
    ''' </summary>
    ''' <remarks>This property may return null for some tables such as those inside the header or footer.</remarks>
    Public ReadOnly Property Table() As Table
        Get
            Return If(Row IsNot Nothing, Row.ParentTable, Nothing)
        End Get
    End Property

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    ''' <remarks>This property may return null for nodes that are inside the header or footer.</remarks>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Dim para As Paragraph
            para = If(Cells.First.Lines.First IsNot Nothing, Cells.First.Lines.First.Paragraph, Nothing)
            Return If(para IsNot Nothing, para.GetAncestor(NodeType.Row), para)
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

    ''' <summary>
    ''' Returns the body that corresponds to the layout entity.  
    ''' </summary>
    Public ReadOnly Property Body() As Body
        Get
            Return CType(ParentNode, Body)
        End Get
    End Property

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Return GetChildEntities(LayoutEntityType.Line, True).First.ParentNode.GetAncestor(NodeType.Body)
        End Get
    End Property
End Class

''' <summary>
''' Represents a table cell.
''' </summary>
Public Class RenderedCell
    Inherits StoryLayoutEntity
    ''' <summary>
    ''' Returns the cell that corresponds to the layout entity.  
    ''' </summary>
    ''' <remarks>This property may return null for some cells such as those inside the header or footer.</remarks>
    Public ReadOnly Property Cell() As Cell
        Get
            Return CType(ParentNode, Cell)
        End Get
    End Property

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    ''' <remarks>This property may return null for nodes that are inside the header or footer.</remarks>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            If Lines.First Is Nothing Then
                Return Nothing
            Else
                Return If(Lines.First.Paragraph IsNot Nothing, Lines.First.Paragraph.GetAncestor(NodeType.Cell), Nothing)
            End If
        End Get
    End Property
End Class

''' <summary>
''' Represents placeholder for footnote content.
''' </summary>
Public Class RenderedFootnote
    Inherits StoryLayoutEntity
    ''' <summary>
    ''' Returns the footnote that corresponds to the layout entity.  
    ''' </summary>
    Public ReadOnly Property Footnote() As Footnote
        Get
            Return CType(ParentNode, Footnote)
        End Get
    End Property

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Return GetChildEntities(LayoutEntityType.Line, True).First.ParentNode.GetAncestor(NodeType.Footnote)
        End Get
    End Property
End Class

''' <summary>
''' Represents placeholder for endnote content.
''' </summary>
Public Class RenderedEndnote
    Inherits StoryLayoutEntity
    ''' <summary>
    ''' Returns the endnote that corresponds to the layout entity.  
    ''' </summary>
    Public ReadOnly Property Endnote() As Footnote
        Get
            Return CType(ParentNode, Footnote)
        End Get
    End Property

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Return GetChildEntities(LayoutEntityType.Line, True).First.ParentNode.GetAncestor(NodeType.Footnote)
        End Get
    End Property
End Class

''' <summary>
''' Represents text area inside of a shape.
''' </summary>
Public Class RenderedTextBox
    Inherits StoryLayoutEntity
    ''' <summary>
    ''' Returns the Shape or DrawingML that corresponds to the layout entity.  
    ''' </summary>
    ''' <remarks>This property may return null for some Shapes or DrawingML such as those inside the header or footer.</remarks>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Dim lines As LayoutCollection(Of LayoutEntity) = GetChildEntities(LayoutEntityType.Line, True)
            Dim shape As Node = lines.First.ParentNode.GetAncestor(NodeType.Shape)

            If shape IsNot Nothing Then
                Return shape
            Else
                Return lines.First.ParentNode.GetAncestor(NodeType.Shape)
            End If
        End Get
    End Property
End Class

''' <summary>
''' Represents placeholder for comment content.
''' </summary>
Public Class RenderedComment
    Inherits StoryLayoutEntity
    ''' <summary>
    ''' Returns the comment that corresponds to the layout entity.  
    ''' </summary>
    Public ReadOnly Property Comment() As Comment
        Get
            Return CType(ParentNode, Comment)
        End Get
    End Property

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Return GetChildEntities(LayoutEntityType.Line, True).First.ParentNode.GetAncestor(NodeType.Comment)
        End Get
    End Property
End Class

''' <summary>
''' Represents footnote/endnote separator.
''' </summary>
Public Class RenderedNoteSeparator
    Inherits StoryLayoutEntity
    ''' <summary>
    ''' Returns the footnote/endnote that corresponds to the layout entity.  
    ''' </summary>
    Public ReadOnly Property Footnote() As Footnote
        Get
            Return CType(ParentNode, Footnote)
        End Get
    End Property

    ''' <summary>
    ''' Returns the node that corresponds to this layout entity.  
    ''' </summary>
    Public Overrides ReadOnly Property ParentNode() As Node
        Get
            Return GetChildEntities(LayoutEntityType.Line, True).First.ParentNode.GetAncestor(NodeType.Footnote)
        End Get
    End Property
End Class

''' <summary>
''' Provides an API wrapper for the LayoutEnumerator class to access the page layout entities of a document presented in
''' a object model like design.
''' </summary>
Public Class RenderedDocument
    Inherits LayoutEntity
    
    Public Sub New(ByVal doc As Document)
        mLayoutCollector = New LayoutCollector(doc)
        mEnumerator = New LayoutEnumerator(doc)
        ProcessLayoutElements(Me)
        LinkLayoutMarkersToNodes(doc)
        CollectLinesAndAddToMarkers()
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
                        If mLayoutToNodeLookup.ContainsKey(span.LayoutObject) Then
                            If span.Kind = "PARAGRAPH" OrElse span.Kind = "ROW" OrElse span.Kind = "CELL" OrElse span.Kind = "SECTION" Then
                                Dim node As Node = mLayoutToNodeLookup(span.LayoutObject)

                                If node.NodeType = NodeType.Row Then
                                    node = (CType(node, Row)).LastCell.LastParagraph
                                End If

                                For Each collectedLine As RenderedLine In collectedLines
                                    collectedLine.SetParentNode(node)
                                Next collectedLine

                                collectedLines = New List(Of RenderedLine)()
                            Else
                                span.SetParentNode(mLayoutToNodeLookup(span.LayoutObject))
                            End If
                        End If
                    Next span
                Next line
            Next story
        Next page
    End Sub

    Private Sub LinkLayoutMarkersToNodes(ByVal doc As Document)
        For Each node As Node In doc.GetChildNodes(NodeType.Any, True)
            Dim entity As Object = mLayoutCollector.GetEntity(node)

            If entity IsNot Nothing Then
                mLayoutToNodeLookup.Add(entity, node)
            End If
        Next node
    End Sub

    Private mLayoutCollector As LayoutCollector
    Private mEnumerator As LayoutEnumerator
    Private Shared mLayoutToNodeLookup As Dictionary(Of Object, Node) = New Dictionary(Of Object, Node)()
End Class
