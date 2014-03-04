' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
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
	''' Splits a document into multiple documents, one per page.
	''' </summary>
	Public Class DocumentPageSplitter
		''' <summary>
		''' Initializes new instance of this class. This method splits the document into sections so that each page 
		''' begins and ends at a section boundary. It is recommended not to modify the document afterwards.
		''' </summary>
		''' <param name="collector">A collector instance which has layout model records for the document.</param>
		Public Sub New(ByVal collector As LayoutCollector)
			mPageNumberFinder = New PageNumberFinder(collector)
			mPageNumberFinder.SplitNodesAcrossPages()
		End Sub

		''' <summary>
		''' Gets the document of a page.
		''' </summary>
		''' <param name="pageIndex">1-based index of a page.</param>
		Public Function GetDocumentOfPage(ByVal pageIndex As Integer) As Document
			Return GetDocumentOfPageRange(pageIndex, pageIndex)
		End Function

		''' <summary>
		''' Gets the document of a page range.
		''' </summary>
		''' <param name="startIndex">1-based index of the start page.</param>
		''' <param name="endIndex">1-based index of the end page.</param>
		Public Function GetDocumentOfPageRange(ByVal startIndex As Integer, ByVal endIndex As Integer) As Document
			Dim result As Document = CType(Document.Clone(False), Document)

			For Each section As Section In mPageNumberFinder.RetrieveAllNodesOnPages(startIndex, endIndex, NodeType.Section)
				result.AppendChild(result.ImportNode(section, True))
			Next section

			Return result
		End Function

		''' <summary>
		''' Gets the document this instance works with.
		''' </summary>
		Private ReadOnly Property Document() As Document
			Get
				Return mPageNumberFinder.Document
			End Get
		End Property

		Private mPageNumberFinder As PageNumberFinder
	End Class
End Namespace
