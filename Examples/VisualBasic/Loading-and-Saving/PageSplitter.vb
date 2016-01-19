
Imports System.Collections
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Tables
Imports System.Diagnostics
Imports Aspose.Words.MailMerging
Imports Aspose.Words.Saving
Imports System.Text
Imports Aspose.Words.Layout
Imports Aspose.Words.Fields
Imports Aspose.Words.Markup
Imports Aspose.Words.Lists

Public Class PageSplitter
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving() + "Split"

        SplitAllDocumentsToPages(dataDir)

        Console.WriteLine(vbNewLine + "Document split to pages successfully." + vbNewLine + "File saved at " + dataDir + "\\Out")
    End Sub

    Public Shared Sub SplitDocumentToPages(ByVal docName As String)
        Dim folderName As String = Path.GetDirectoryName(docName)
        Dim fileName As String = Path.GetFileNameWithoutExtension(docName)
        Dim extensionName As String = Path.GetExtension(docName)
        Dim outFolder As String = Path.Combine(folderName, "Out")

        Console.WriteLine("Processing document: " & fileName & extensionName)

        Dim doc As New Document(docName)

        ' Create and attach collector to the document before page layout is built.
        Dim layoutCollector As New LayoutCollector(doc)

        ' This will build layout model and collect necessary information.
        doc.UpdatePageLayout()

        ' Split nodes in the document into separate pages.
        Dim splitter As New DocumentPageSplitter(layoutCollector)

        ' Save each page to the disk as a separate document.
        For page As Integer = 1 To doc.PageCount
            Dim pageDoc As Document = splitter.GetDocumentOfPage(page)
            pageDoc.Save(Path.Combine(outFolder, String.Format("{0} - page{1} Out{2}", fileName, page, extensionName)))
        Next page

        ' Detach the collector from the document.
        layoutCollector.Document = Nothing
    End Sub

    Public Shared Sub SplitAllDocumentsToPages(ByVal folderName As String)
        Dim fileNames() As String = Directory.GetFiles(folderName, "*.doc?", SearchOption.TopDirectoryOnly)

        For Each fileName As String In fileNames
            SplitDocumentToPages(fileName)
        Next fileName
    End Sub
End Class


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