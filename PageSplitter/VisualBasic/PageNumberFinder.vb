'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.Layout

Namespace PageSplitter
	''' <summary>
	''' Provides methods for extracting nodes of a document which are rendered on a specified pages.
	''' </summary>
	Public Class PageNumberFinder
		''' <summary>
		''' Initializes new instance of this class.
		''' </summary>
		''' <param name="collector">A collector instance which has layout model records for the document.</param>
		Public Sub New(ByVal collector As LayoutCollector)
			mCollector = collector
		End Sub

		''' <summary>
		''' Retrieves 1-based index of a page that the node begins on.
		''' </summary>
		Public Function GetPage(ByVal node As Node) As Integer
			If mNodeStartPageLookup.ContainsKey(node) Then
			   Return CInt(Fix(mNodeStartPageLookup(node)))
			End If

			Return mCollector.GetStartPageIndex(node)
		End Function

		''' <summary>
		''' Retrieves 1-based index of a page that the node ends on.
		''' </summary>
		Public Function GetPageEnd(ByVal node As Node) As Integer
			If mNodeEndPageLookup.ContainsKey(node) Then
				Return CInt(Fix(mNodeEndPageLookup(node)))
			End If

			Return mCollector.GetEndPageIndex(node)
		End Function

		''' <summary>
		''' Returns how many pages the specified node spans over. Returns 1 if the node is contained within one page.
		''' </summary>
		Public Function PageSpan(ByVal node As Node) As Integer
			Return GetPageEnd(node) - GetPage(node) + 1
		End Function

		''' <summary>
		''' Returns a list of nodes that are contained anywhere on the specified page or pages which match the specified node type.
		''' </summary>
		Public Function RetrieveAllNodesOnPages(ByVal startPage As Integer, ByVal endPage As Integer, ByVal nodeType As NodeType) As ArrayList
			If startPage < 1 OrElse startPage > Document.PageCount Then
				Throw New ArgumentOutOfRangeException("startPage")
			End If

			If endPage < 1 OrElse endPage > Document.PageCount OrElse endPage < startPage Then
				Throw New ArgumentOutOfRangeException("endPage")
			End If

			CheckPageListsPopulated()

			Dim pageNodes As New ArrayList()

			For page As Integer = startPage To endPage
				' Some pages can be empty.
				If (Not mReversePageLookup.ContainsKey(page)) Then
					Continue For
				End If

				For Each node As Node In CType(mReversePageLookup(page), ArrayList)
					If node.ParentNode IsNot Nothing AndAlso (nodeType = NodeType.Any OrElse node.NodeType = nodeType) AndAlso (Not pageNodes.Contains(node)) Then
						pageNodes.Add(node)
					End If
				Next node
			Next page

			Return pageNodes
		End Function

		''' <summary>
		''' Splits nodes which appear over two or more pages into separate nodes so that they still appear in the same way
		''' but no longer appear across a page.
		''' </summary>
		Public Sub SplitNodesAcrossPages()
			' Visit any composites which are possibly split across pages and split them into separate nodes.
			Document.Accept(New SectionSplitter(Me))
		End Sub

		''' <summary>
		''' Gets the document this instance works with.
		''' </summary>
		Public ReadOnly Property Document() As Document
			Get
				Return mCollector.Document
			End Get
		End Property

		''' <summary>
		''' This is called by <see cref="SectionSplitter"/> to update page numbers of split nodes.
		''' </summary>
		Friend Sub AddPageNumbersForNode(ByVal node As Node, ByVal startPage As Integer, ByVal endPage As Integer)
			If startPage > 0 Then
				mNodeStartPageLookup(node) = startPage
			End If

			If endPage > 0 Then
				mNodeEndPageLookup(node) = endPage
			End If
		End Sub

		Private Sub CheckPageListsPopulated()
			If mReversePageLookup IsNot Nothing Then
				Return
			End If

			mReversePageLookup = New Hashtable()

			' Add each node to a list which represent the nodes found on each page.
			For Each node As Node In Document.GetChildNodes(NodeType.Any, True)
				' Headers/Footers follow sections. They are not split by themselves.
				If IsHeaderFooterType(node) Then
					Continue For
				End If

				Dim startPage As Integer = GetPage(node)
				Dim endPage As Integer = GetPageEnd(node)

				For page As Integer = startPage To endPage
					If (Not mReversePageLookup.ContainsKey(page)) Then
						mReversePageLookup.Add(page, New ArrayList())
					End If

					CType(mReversePageLookup(page), ArrayList).Add(node)
				Next page
			Next node
		End Sub

		Private Shared Function IsHeaderFooterType(ByVal node As Node) As Boolean
			Return node.NodeType = NodeType.HeaderFooter OrElse node.GetAncestor(NodeType.HeaderFooter) IsNot Nothing
		End Function

		' Maps node to a start/end page numbers. This is used to override baseline page numbers provided by collector when document is split.
		Private mNodeStartPageLookup As New Hashtable()
		Private mNodeEndPageLookup As New Hashtable()
		' Maps page number to a list of nodes found on that page.
		Private mReversePageLookup As Hashtable
		Private mCollector As LayoutCollector
	End Class
End Namespace