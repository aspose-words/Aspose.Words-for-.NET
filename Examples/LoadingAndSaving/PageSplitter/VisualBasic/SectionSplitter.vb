'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
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
Imports Aspose.Words.Fields
Imports Aspose.Words.Lists

Namespace PageSplitterExample
	''' <summary>
	''' Splits a document into multiple sections so that each page begins and ends at a section boundary.
	''' </summary>
	Public Class SectionSplitter
		Inherits DocumentVisitor
		Public Sub New(ByVal pageNumberFinder As PageNumberFinder)
			mPageNumberFinder = pageNumberFinder
		End Sub

		Public Overrides Function VisitParagraphStart(ByVal paragraph As Paragraph) As VisitorAction
			If paragraph.IsListItem Then
				Dim paraList As List = paragraph.ListFormat.List
				Dim currentLevel As ListLevel = paragraph.ListFormat.ListLevel

				' Since we have encountered a list item we need to check if this will reset
				' any subsequent list levels and if so then update the numbering of the level.
				Dim currentListLevelNumber As Integer = paragraph.ListFormat.ListLevelNumber
				For i As Integer = currentListLevelNumber + 1 To paraList.ListLevels.Count - 1
					Dim paraLevel As ListLevel = paraList.ListLevels(i)

					If paraLevel.RestartAfterLevel >= currentListLevelNumber Then
						' This list level needs to be reset after the current list number.
						mListLevelToListNumberLookup(paraLevel) = paraLevel.StartAt
					End If
				Next i

				' A list which was used on a previous page is present on a different page, the list
				' needs to be copied so list numbering is retained when extracting individual pages.
				If ContainsListLevelAndPageChanged(paragraph) Then
					Dim copyList As List = paragraph.Document.Lists.AddCopy(paraList)
					mListLevelToListNumberLookup(currentLevel) = paragraph.ListLabel.LabelValue

					' Set the numbering of each list level to start at the numbering of the level on the previous page.
					For i As Integer = 0 To paraList.ListLevels.Count - 1
						Dim paraLevel As ListLevel = paraList.ListLevels(i)

						If mListLevelToListNumberLookup.ContainsKey(paraLevel) Then
							copyList.ListLevels(i).StartAt = CInt(Fix(mListLevelToListNumberLookup(paraLevel)))
						End If
					Next i

					mListToReplacementListLookup(paraList) = copyList
				End If

				If mListToReplacementListLookup.ContainsKey(paraList) Then
					' This paragraph belongs to a list from a previous page. Apply the replacement list.
					paragraph.ListFormat.List = CType(mListToReplacementListLookup(paraList), List)
					' This is a trick to get the spacing of the list level to set correctly.
					paragraph.ListFormat.ListLevelNumber += 0
				End If

				mListLevelToPageLookup(currentLevel) = mPageNumberFinder.GetPage(paragraph)
				mListLevelToListNumberLookup(currentLevel) = paragraph.ListLabel.LabelValue
			End If

			Dim prevSection As Section = CType(paragraph.ParentSection.PreviousSibling, Section)
			Dim prevBodyPara As Paragraph = TryCast(paragraph.PreviousSibling, Paragraph)

			Dim prevSectionPara As Paragraph = If(prevSection IsNot Nothing AndAlso paragraph Is paragraph.ParentSection.Body.FirstChild, prevSection.Body.LastParagraph, Nothing)
			Dim prevParagraph As Paragraph = If(prevBodyPara IsNot Nothing, prevBodyPara, prevSectionPara)

			If paragraph.IsEndOfSection AndAlso (Not paragraph.HasChildNodes) Then
				paragraph.Remove()
			End If

			' Paragraphs across pages can merge or remove spacing depending upon the previous paragraph.
			If prevParagraph IsNot Nothing Then
				If mPageNumberFinder.GetPage(paragraph) <> mPageNumberFinder.GetPageEnd(prevParagraph) Then
					If paragraph.IsListItem AndAlso prevParagraph.IsListItem AndAlso (Not prevParagraph.IsEndOfSection) Then
					   prevParagraph.ParagraphFormat.SpaceAfter = 0
					ElseIf prevParagraph.ParagraphFormat.StyleName = paragraph.ParagraphFormat.StyleName AndAlso paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle Then
						paragraph.ParagraphFormat.SpaceBefore = 0
					ElseIf paragraph.ParagraphFormat.PageBreakBefore OrElse (prevParagraph.IsEndOfSection AndAlso prevSection.PageSetup.SectionStart <> SectionStart.NewColumn) Then
						paragraph.ParagraphFormat.SpaceBefore = System.Math.Max(paragraph.ParagraphFormat.SpaceBefore - prevParagraph.ParagraphFormat.SpaceAfter, 0)
					Else
						paragraph.ParagraphFormat.SpaceBefore = 0
					End If
				End If
			End If

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitSectionStart(ByVal section As Section) As VisitorAction
			mSectionCount += 1
			Dim previousSection As Section = CType(section.PreviousSibling, Section)

			' If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an 
			' extracted document if the previous section is missing.
			If previousSection IsNot Nothing Then
				If (Not section.PageSetup.RestartPageNumbering) Then
					section.PageSetup.RestartPageNumbering = True
					section.PageSetup.PageStartingNumber = previousSection.PageSetup.PageStartingNumber + mPageNumberFinder.PageSpan(previousSection)
				End If

				For Each previousHeaderFooter As HeaderFooter In previousSection.HeadersFooters
					If section.HeadersFooters(previousHeaderFooter.HeaderFooterType) Is Nothing Then
						Dim newHeaderFooter As HeaderFooter = CType(previousSection.HeadersFooters(previousHeaderFooter.HeaderFooterType).Clone(True), HeaderFooter)
						section.HeadersFooters.Add(newHeaderFooter)
					End If
				Next previousHeaderFooter
			End If

			' Manually set the result of these fields before sections are split.
			For Each headerFooter As HeaderFooter In section.HeadersFooters
				For Each field As Field In headerFooter.Range.Fields
					If field.Type = FieldType.FieldSection OrElse field.Type = FieldType.FieldSectionPages Then
						field.Result = If((field.Type = FieldType.FieldSection), mSectionCount.ToString(), mPageNumberFinder.PageSpan(section).ToString())
						field.IsLocked = True
					End If
				Next field
			Next headerFooter

			' All fields in the body should stay the same, this also improves field update time.
			For Each field As Field In section.Body.Range.Fields
				field.IsLocked = True
			Next field

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitDocumentEnd(ByVal doc As Document) As VisitorAction
			' All sections have separate headers and footers now, update the fields in all headers and footers
			' to the correct values. This allows each page to maintain the correct field results even when
			' PAGE or IF fields are used.
			doc.UpdateFields()

			For Each headerFooter As HeaderFooter In doc.GetChildNodes(NodeType.HeaderFooter, True)
				For Each field As Field In headerFooter.Range.Fields
					field.IsLocked = True
				Next field
			Next headerFooter

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitSmartTagEnd(ByVal smartTag As SmartTag) As VisitorAction
			If IsCompositeAcrossPage(smartTag) Then
				SplitComposite(smartTag)
			End If

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitCustomXmlMarkupEnd(ByVal customXmlMarkup As CustomXmlMarkup) As VisitorAction
			If IsCompositeAcrossPage(customXmlMarkup) Then
				SplitComposite(customXmlMarkup)
			End If

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitStructuredDocumentTagEnd(ByVal sdt As StructuredDocumentTag) As VisitorAction
			If IsCompositeAcrossPage(sdt) Then
				SplitComposite(sdt)
			End If

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitCellEnd(ByVal cell As Cell) As VisitorAction
			If IsCompositeAcrossPage(cell) Then
				SplitComposite(cell)
			End If

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitRowEnd(ByVal row As Row) As VisitorAction
			If IsCompositeAcrossPage(row) Then
				SplitComposite(row)
			End If

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitTableEnd(ByVal table As Table) As VisitorAction
			If IsCompositeAcrossPage(table) Then
				' Copy any header rows to other pages.
				Dim stack As New Stack(table.Rows.ToArray())

				For Each cloneTable As Table In SplitComposite(table)
					For Each row As Row In stack
						If row.RowFormat.HeadingFormat Then
							cloneTable.PrependChild(row.Clone(True))
						End If
					Next row
				Next cloneTable
			End If

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitParagraphEnd(ByVal paragraph As Paragraph) As VisitorAction
			If IsCompositeAcrossPage(paragraph) Then
				For Each clonePara As Paragraph In SplitComposite(paragraph)
					' Remove list numbering from the cloned paragraph but leave the indent the same 
					' as the paragraph is supposed to be part of the item before.
					If paragraph.IsListItem Then
						Dim textPosition As Double = clonePara.ListFormat.ListLevel.TextPosition
						clonePara.ListFormat.RemoveNumbers()
						clonePara.ParagraphFormat.LeftIndent = textPosition
					End If

					' Reset spacing of split paragraphs as additional spacing is removed.
					clonePara.ParagraphFormat.SpaceBefore = 0
					paragraph.ParagraphFormat.SpaceAfter = 0
				Next clonePara
			End If

			Return VisitorAction.Continue
		End Function

		Public Overrides Function VisitSectionEnd(ByVal section As Section) As VisitorAction
			If IsCompositeAcrossPage(section) Then
				' If a TOC field spans across more than one page then the hyperlink formatting may show through.
				' Remove direct formatting to avoid this.
				For Each start As FieldStart In section.GetChildNodes(NodeType.FieldStart, True)
					If start.FieldType = FieldType.FieldTOC Then
						Dim field As Field = start.GetField()
						Dim node As Node = field.Separator

						node = node.NextPreOrder(section)
						Do While node IsNot field.End
							If node.NodeType = NodeType.Run Then
								CType(node, Run).Font.ClearFormatting()
							End If
							node = node.NextPreOrder(section)
						Loop
					End If
				Next start

				For Each cloneSection As Section In SplitComposite(section)
					cloneSection.PageSetup.SectionStart = SectionStart.NewPage
					cloneSection.PageSetup.RestartPageNumbering = True
					cloneSection.PageSetup.PageStartingNumber = section.PageSetup.PageStartingNumber + (section.Document.IndexOf(cloneSection) - section.Document.IndexOf(section))
					cloneSection.PageSetup.DifferentFirstPageHeaderFooter = False

					RemovePageBreaksFromParagraph(cloneSection.Body.LastParagraph)
				Next cloneSection

				RemovePageBreaksFromParagraph(section.Body.LastParagraph)

				' Add new page numbering for the body of the section as well.
				mPageNumberFinder.AddPageNumbersForNode(section.Body, mPageNumberFinder.GetPage(section), mPageNumberFinder.GetPageEnd(section))
			End If

			Return VisitorAction.Continue
		End Function

		Private Function IsCompositeAcrossPage(ByVal composite As CompositeNode) As Boolean
			Return mPageNumberFinder.PageSpan(composite) > 1
		End Function

		Private Function ContainsListLevelAndPageChanged(ByVal para As Paragraph) As Boolean
			Return mListLevelToPageLookup.ContainsKey(para.ListFormat.ListLevel) AndAlso CInt(Fix(mListLevelToPageLookup(para.ListFormat.ListLevel))) <> mPageNumberFinder.GetPage(para)
		End Function

		Private Sub RemovePageBreaksFromParagraph(ByVal para As Paragraph)
			If para IsNot Nothing Then
				For Each run As Run In para.Runs
					run.Text = run.Text.Replace(ControlChar.PageBreak, String.Empty)
				Next run
			End If
		End Sub

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

		Private mListLevelToListNumberLookup As New Hashtable()
		Private mListToReplacementListLookup As New Hashtable()
		Private mListLevelToPageLookup As New Hashtable()
		Private mPageNumberFinder As PageNumberFinder
		Private mSectionCount As Integer
	End Class
End Namespace