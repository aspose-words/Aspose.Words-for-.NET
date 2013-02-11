'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.Tables
Imports Aspose.Words.Markup

Namespace PageSplitter
	''' <summary>
	''' Splits a document into multiple sections so that each page begins and ends at a section boundary.
	''' </summary>
	Public Class SectionSplitter
		Inherits DocumentVisitor
		Public Sub New(ByVal pageNumberFinder As PageNumberFinder)
			mPageNumberFinder = pageNumberFinder
		End Sub

		Public Overrides Function VisitParagraphStart(ByVal paragraph As Paragraph) As VisitorAction
			Return ContinueIfCompositeAcrossPageElseSkip(paragraph)
		End Function

		Public Overrides Function VisitTableStart(ByVal table As Table) As VisitorAction
			Return ContinueIfCompositeAcrossPageElseSkip(table)
		End Function

		Public Overrides Function VisitRowStart(ByVal row As Row) As VisitorAction
			Return ContinueIfCompositeAcrossPageElseSkip(row)
		End Function

		Public Overrides Function VisitCellStart(ByVal cell As Cell) As VisitorAction
			Return ContinueIfCompositeAcrossPageElseSkip(cell)
		End Function

		Public Overrides Function VisitCustomXmlMarkupStart(ByVal customXmlMarkup As CustomXmlMarkup) As VisitorAction
			Return ContinueIfCompositeAcrossPageElseSkip(customXmlMarkup)
		End Function

		Public Overrides Function VisitStructuredDocumentTagStart(ByVal sdt As StructuredDocumentTag) As VisitorAction
			Return ContinueIfCompositeAcrossPageElseSkip(sdt)
		End Function

		Public Overrides Function VisitSmartTagStart(ByVal smartTag As SmartTag) As VisitorAction
			Return ContinueIfCompositeAcrossPageElseSkip(smartTag)
		End Function

		Public Overrides Function VisitSectionStart(ByVal section As Section) As VisitorAction
			Dim previousSection As Section = CType(section.PreviousSibling, Section)

			' If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an 
			' extracted document if the previous section is missing.
			If previousSection IsNot Nothing Then
				Dim previousHeaderFooters As HeaderFooterCollection = previousSection.HeadersFooters

				If (Not section.PageSetup.RestartPageNumbering) Then
					section.PageSetup.RestartPageNumbering = True
					section.PageSetup.PageStartingNumber = previousSection.PageSetup.PageStartingNumber + mPageNumberFinder.PageSpan(previousSection)
				End If

				For Each previousHeaderFooter As HeaderFooter In previousHeaderFooters
					If section.HeadersFooters(previousHeaderFooter.HeaderFooterType) Is Nothing Then
						Dim newHeaderFooter As HeaderFooter = CType(previousHeaderFooters(previousHeaderFooter.HeaderFooterType).Clone(True), HeaderFooter)
						section.HeadersFooters.Add(newHeaderFooter)
					End If
				Next previousHeaderFooter
			End If

			Return ContinueIfCompositeAcrossPageElseSkip(section)
		End Function

		Public Overrides Function VisitSmartTagEnd(ByVal smartTag As SmartTag) As VisitorAction
			SplitComposite(smartTag)
			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitCustomXmlMarkupEnd(ByVal customXmlMarkup As CustomXmlMarkup) As VisitorAction
			SplitComposite(customXmlMarkup)
			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitStructuredDocumentTagEnd(ByVal sdt As StructuredDocumentTag) As VisitorAction
			SplitComposite(sdt)
			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitCellEnd(ByVal cell As Cell) As VisitorAction
			SplitComposite(cell)
			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitRowEnd(ByVal row As Row) As VisitorAction
			SplitComposite(row)
			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitTableEnd(ByVal table As Table) As VisitorAction
			SplitComposite(table)
			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitParagraphEnd(ByVal paragraph As Paragraph) As VisitorAction
			For Each clonePara As Paragraph In SplitComposite(paragraph)
				' Remove list numbering from the cloned paragraph but leave the indent the same 
				' as the paragraph is supposed to be part of the item before.
				If paragraph.IsListItem Then
					Dim textPosition As Double = clonePara.ListFormat.ListLevel.TextPosition
					clonePara.ListFormat.RemoveNumbers()
					clonePara.ParagraphFormat.LeftIndent = textPosition
				End If

				' Reset spacing of split paragraphs in tables as additional spacing may cause them to look different.
				If paragraph.IsInCell Then
					clonePara.ParagraphFormat.SpaceBefore = 0
					paragraph.ParagraphFormat.SpaceAfter = 0
				End If
			Next clonePara

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitSectionEnd(ByVal section As Section) As VisitorAction
			For Each cloneSection As Section In SplitComposite(section)
				cloneSection.PageSetup.SectionStart = SectionStart.NewPage
				cloneSection.PageSetup.RestartPageNumbering = True
				cloneSection.PageSetup.PageStartingNumber = section.PageSetup.PageStartingNumber + (section.Document.IndexOf(cloneSection) - section.Document.IndexOf(section))
				cloneSection.PageSetup.DifferentFirstPageHeaderFooter = False
			Next cloneSection

			' Add new page numbering for the body of the section as well.
			mPageNumberFinder.AddPageNumbersForNode(section.Body, mPageNumberFinder.GetPage(section), mPageNumberFinder.GetPageEnd(section))

			Return VisitorAction.Continue
		End Function

		Private Function ContinueIfCompositeAcrossPageElseSkip(ByVal composite As CompositeNode) As VisitorAction
			Return If((mPageNumberFinder.PageSpan(composite) > 1), VisitorAction.Continue, VisitorAction.SkipThisNode)
		End Function

		Private Function SplitComposite(ByVal composite As CompositeNode) As ArrayList
			Dim splitNodes As New ArrayList()
			For Each splitNode As Node In FindChildSplitPositions(composite)
				splitNodes.Add(SplitCompositeAtNode(composite, splitNode))
			Next splitNode

			Return splitNodes
		End Function

		Private Function FindChildSplitPositions(ByVal node As CompositeNode) As ArrayList
			' A node may span across multiple pages so a list of split positions is returned.
			' The split node is the first node on the next page.
			Dim splitList As New ArrayList()

			Dim startingPage As Integer = mPageNumberFinder.GetPage(node)

			Dim childNodes() As Node = If(node.NodeType = NodeType.Section, (CType(node, Section)).Body.ChildNodes.ToArray(), node.ChildNodes.ToArray())

			For Each childNode As Node In childNodes
				Dim pageNum As Integer = mPageNumberFinder.GetPage(childNode)

				' If the page of the child node has changed then this is the split position. Add
				' this to the list.
				If pageNum > startingPage Then
					splitList.Add(childNode)
					startingPage = pageNum
				End If

				If mPageNumberFinder.PageSpan(childNode) > 1 Then
					mPageNumberFinder.AddPageNumbersForNode(childNode, pageNum, pageNum)
				End If
			Next childNode

			' Split composites backward so the cloned nodes are inserted in the right order.
			splitList.Reverse()

			Return splitList
		End Function

		Private Function SplitCompositeAtNode(ByVal baseNode As CompositeNode, ByVal targetNode As Node) As CompositeNode
			Dim cloneNode As CompositeNode = CType(baseNode.Clone(False), CompositeNode)

			Dim node As Node = targetNode
			Dim currentPageNum As Integer = mPageNumberFinder.GetPage(baseNode)

			' Move all nodes found on the next page into the copied node. Handle row nodes separately.
			If baseNode.NodeType <> NodeType.Row Then
				Dim composite As CompositeNode = cloneNode

				If baseNode.NodeType = NodeType.Section Then
					cloneNode = CType(baseNode.Clone(True), CompositeNode)
					Dim section As Section = CType(cloneNode, Section)
					section.Body.RemoveAllChildren()

					composite = section.Body
				End If

				Do While node IsNot Nothing
					Dim nextNode As Node = node.NextSibling
					composite.AppendChild(node)
					node = nextNode
				Loop
			Else
				' If we are dealing with a row then we need to add in dummy cells for the cloned row.
				Dim targetPageNum As Integer = mPageNumberFinder.GetPage(targetNode)
				Dim childNodes() As Node = baseNode.ChildNodes.ToArray()

				For Each childNode As Node In childNodes
					Dim pageNum As Integer = mPageNumberFinder.GetPage(childNode)

					If pageNum = targetPageNum Then
						cloneNode.LastChild.Remove()
						cloneNode.AppendChild(childNode)
					ElseIf pageNum = currentPageNum Then
						cloneNode.AppendChild(childNode.Clone(False))
						If cloneNode.LastChild.NodeType <> NodeType.Cell Then
							CType(cloneNode.LastChild, CompositeNode).AppendChild((CType(childNode, CompositeNode)).FirstChild.Clone(False))
						End If
					End If
				Next childNode
			End If

			' Insert the split node after the original.
			baseNode.ParentNode.InsertAfter(cloneNode, baseNode)

			' Update the new page numbers of the base node and the clone node including its descendents.
			' This will only be a single page as the cloned composite is split to be on one page.
			Dim currentEndPageNum As Integer = mPageNumberFinder.GetPageEnd(baseNode)
			mPageNumberFinder.AddPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1)
			mPageNumberFinder.AddPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum)

			For Each childNode As Node In cloneNode.GetChildNodes(NodeType.Any, True)
				mPageNumberFinder.AddPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum)
			Next childNode

			Return cloneNode
		End Function

		Private mPageNumberFinder As PageNumberFinder
	End Class
End Namespace