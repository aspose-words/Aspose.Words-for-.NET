'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Reflection
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.Tables
Imports Aspose.Words.Fields

Public Class ExtractContent

    ' The path to the documents directory.
    Shared dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()

    Public Shared Sub Run()

        ' Call methods to test extraction of different types from the document.
        ExtractContentBetweenParagraphs()
        ExtractContentBetweenBlockLevelNodes()
        ExtractContentBetweenParagraphStyles()
        ExtractContentBetweenRuns()
        ExtractContentUsingField()
        ExtractContentBetweenBookmark()
        ExtractContentBetweenCommentRange()

        Console.WriteLine(vbNewLine & "Comments extracted and removed successfully." & vbNewLine & "File saved at " + dataDir + "Test File Out.doc")
    End Sub

    Public Shared Sub ExtractContentBetweenParagraphs()
        ' Load in the document
        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Gather the nodes. The GetChild method uses 0-based index
        Dim startPara As Paragraph = CType(doc.FirstSection.GetChild(NodeType.Paragraph, 6, True), Paragraph)
        Dim endPara As Paragraph = CType(doc.FirstSection.GetChild(NodeType.Paragraph, 10, True), Paragraph)
        ' Extract the content between these nodes in the document. Include these markers in the extraction.
        Dim extractedNodes As ArrayList = ExtractContent(startPara, endPara, True)

        ' Insert the content into a new separate document and save it to disk.
        Dim dstDoc As Document = GenerateDocument(doc, extractedNodes)
        dstDoc.Save(dataDir & "TestFile.Paragraphs Out.doc")

        Console.WriteLine(vbNewLine & "Extracted content betweenn the paragraphs successfully." & vbNewLine & "File saved at " + dataDir + "TestFile.Paragraphs Out.doc")
    End Sub

    Public Shared Sub ExtractContentBetweenBlockLevelNodes()
        Dim doc As New Document(dataDir & "TestFile.doc")

        Dim startPara As Paragraph = CType(doc.LastSection.GetChild(NodeType.Paragraph, 2, True), Paragraph)
        Dim endTable As Table = CType(doc.LastSection.GetChild(NodeType.Table, 0, True), Table)

        ' Extract the content between these nodes in the document. Include these markers in the extraction.
        Dim extractedNodes As ArrayList = ExtractContent(startPara, endTable, True)

        ' Lets reverse the array to make inserting the content back into the document easier.
        extractedNodes.Reverse()

        Do While extractedNodes.Count > 0
            ' Insert the last node from the reversed list 
            endTable.ParentNode.InsertAfter(CType(extractedNodes(0), Node), endTable)
            ' Remove this node from the list after insertion.
            extractedNodes.RemoveAt(0)
        Loop

        ' Save the generated document to disk.
        doc.Save(dataDir & "TestFile.DuplicatedContent Out.doc")

        Console.WriteLine(vbNewLine & "Extracted content betweenn the block level nodes successfully." & vbNewLine & "File saved at " + dataDir + "TestFile.DuplicatedContent Out.doc")
    End Sub

    Public Shared Sub ExtractContentBetweenParagraphStyles()
        ' Load in the document
        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Gather a list of the paragraphs using the respective heading styles.
        Dim parasStyleHeading1 As ArrayList = ParagraphsByStyleName(doc, "Heading 1")
        Dim parasStyleHeading3 As ArrayList = ParagraphsByStyleName(doc, "Heading 3")

        ' Use the first instance of the paragraphs with those styles.
        Dim startPara1 As Node = CType(parasStyleHeading1(0), Node)
        Dim endPara1 As Node = CType(parasStyleHeading3(0), Node)

        ' Extract the content between these nodes in the document. Don't include these markers in the extraction.
        Dim extractedNodes As ArrayList = ExtractContent(startPara1, endPara1, False)

        ' Insert the content into a new separate document and save it to disk.
        Dim dstDoc As Document = GenerateDocument(doc, extractedNodes)
        dstDoc.Save(dataDir & "TestFile.Styles Out.doc")

        Console.WriteLine(vbNewLine & "Extracted content betweenn the paragraph styles successfully." & vbNewLine & "File saved at " + dataDir + "TestFile.Styles Out.doc")
    End Sub

    Public Shared Sub ExtractContentBetweenRuns()
        ' Load in the document
        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Retrieve a paragraph from the first section.
        Dim para As Paragraph = CType(doc.GetChild(NodeType.Paragraph, 7, True), Paragraph)

        ' Use some runs for extraction.
        Dim startRun As Run = para.Runs(1)
        Dim endRun As Run = para.Runs(4)

        ' Extract the content between these nodes in the document. Include these markers in the extraction.
        Dim extractedNodes As ArrayList = ExtractContent(startRun, endRun, True)

        ' Get the node from the list. There should only be one paragraph returned in the list.
        Dim node As Node = CType(extractedNodes(0), Node)
        ' Print the text of this node to the console.
        Console.WriteLine(node.ToString(SaveFormat.Text))

    End Sub

    Public Shared Sub ExtractContentUsingField()
        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Use a document builder to retrieve the field start of a merge field.
        Dim builder As New DocumentBuilder(doc)

        ' Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        ' We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder.MoveToMergeField("Fullname", False, False)

        ' The builder cursor should be positioned at the start of the field.
        Dim startField As FieldStart = CType(builder.CurrentNode, FieldStart)
        Dim endPara As Paragraph = CType(doc.FirstSection.GetChild(NodeType.Paragraph, 5, True), Paragraph)

        ' Extract the content between these nodes in the document. Don't include these markers in the extraction.
        Dim extractedNodes As ArrayList = ExtractContent(startField, endPara, False)

        ' Insert the content into a new separate document and save it to disk.
        Dim dstDoc As Document = GenerateDocument(doc, extractedNodes)
        dstDoc.Save(dataDir & "TestFile.Fields Out.pdf")

        Console.WriteLine(vbNewLine & "Extracted content using the field successfully." & vbNewLine & "File saved at " + dataDir + "TestFile.Fields Out.pdf")
    End Sub

    Public Shared Sub ExtractContentBetweenBookmark()
        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Retrieve the bookmark from the document.
        Dim bookmark As Aspose.Words.Bookmark = doc.Range.Bookmarks("Bookmark1")

        ' We use the BookmarkStart and BookmarkEnd nodes as markers.
        Dim bookmarkStart As BookmarkStart = bookmark.BookmarkStart
        Dim bookmarkEnd As BookmarkEnd = bookmark.BookmarkEnd

        ' Firstly extract the content between these nodes including the bookmark. 
        Dim extractedNodesInclusive As ArrayList = ExtractContent(bookmarkStart, bookmarkEnd, True)
        Dim dstDoc As Document = GenerateDocument(doc, extractedNodesInclusive)
        dstDoc.Save(dataDir & "TestFile.BookmarkInclusive Out.doc")

        ' Secondly extract the content between these nodes this time without including the bookmark.
        Dim extractedNodesExclusive As ArrayList = ExtractContent(bookmarkStart, bookmarkEnd, False)
        dstDoc = GenerateDocument(doc, extractedNodesExclusive)
        dstDoc.Save(dataDir & "TestFile.BookmarkExclusive Out.doc")

        Console.WriteLine(vbNewLine & "Extracted content between bookmarks successfully." & vbNewLine & "File saved at " + dataDir + "TestFile.BookmarkExclusive Out.doc")
    End Sub

    Public Shared Sub ExtractContentBetweenCommentRange()
        Dim doc As New Document(dataDir & "TestFile.doc")

        ' This is a quick way of getting both comment nodes.
        ' Your code should have a proper method of retrieving each corresponding start and end node.
        Dim commentStart As CommentRangeStart = CType(doc.GetChild(NodeType.CommentRangeStart, 0, True), CommentRangeStart)
        Dim commentEnd As CommentRangeEnd = CType(doc.GetChild(NodeType.CommentRangeEnd, 0, True), CommentRangeEnd)

        ' Firstly extract the content between these nodes including the comment as well. 
        Dim extractedNodesInclusive As ArrayList = ExtractContent(commentStart, commentEnd, True)
        Dim dstDoc As Document = GenerateDocument(doc, extractedNodesInclusive)
        dstDoc.Save(dataDir & "TestFile.CommentInclusive Out.doc")

        ' Secondly extract the content between these nodes without the comment.
        Dim extractedNodesExclusive As ArrayList = ExtractContent(commentStart, commentEnd, False)
        dstDoc = GenerateDocument(doc, extractedNodesExclusive)
        dstDoc.Save(dataDir & "TestFile.CommentExclusive Out.doc")

        Console.WriteLine(vbNewLine & "Extracted content between comment range successfully." & vbNewLine & "File saved at " + dataDir + "TestFile.CommentExclusive Out.doc")
    End Sub

    Public Shared Function ExtractContent(ByVal startNode As Node, ByVal endNode As Node, ByVal isInclusive As Boolean) As ArrayList
        ' First check that the nodes passed to this method are valid for use.
        VerifyParameterNodes(startNode, endNode)

        ' Create a list to store the extracted nodes.
        Dim nodes As New ArrayList()

        ' Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
        Dim originalStartNode As Node = startNode
        Dim originalEndNode As Node = endNode

        ' Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        ' We will split the content of first and last nodes depending if the marker nodes are inline
        Do While startNode.ParentNode.NodeType <> NodeType.Body
            startNode = startNode.ParentNode
        Loop

        Do While endNode.ParentNode.NodeType <> NodeType.Body
            endNode = endNode.ParentNode
        Loop

        Dim isExtracting As Boolean = True
        Dim isStartingNode As Boolean = True
        Dim isEndingNode As Boolean = False
        ' The current node we are extracting from the document.
        Dim currNode As Node = startNode

        ' Begin extracting content. Process all block level nodes and specifically split the first and last nodes when needed so paragraph formatting is retained.
        ' Method is little more complex than a regular extractor as we need to factor in extracting using inline nodes, fields, bookmarks etc as to make it really useful.
        Do While isExtracting
            ' Clone the current node and its children to obtain a copy.
            Dim cloneNode As CompositeNode = CType(currNode.Clone(True), CompositeNode)
            isEndingNode = currNode.Equals(endNode)

            If isStartingNode OrElse isEndingNode Then
                ' We need to process each marker separately so pass it off to a separate method instead.
                If isStartingNode Then
                    ProcessMarker(cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode)
                    isStartingNode = False
                End If

                ' Conditional needs to be separate as the block level start and end markers maybe the same node.
                If isEndingNode Then
                    ProcessMarker(cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode)
                    isExtracting = False
                End If
            Else
                ' Node is not a start or end marker, simply add the copy to the list.
                nodes.Add(cloneNode)
            End If

            ' Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
            If currNode.NextSibling Is Nothing AndAlso isExtracting Then
                ' Move to the next section.
                Dim nextSection As Section = CType(currNode.GetAncestor(NodeType.Section).NextSibling, Section)
                currNode = nextSection.Body.FirstChild
            Else
                ' Move to the next node in the body.
                currNode = currNode.NextSibling
            End If
        Loop

        ' Return the nodes between the node markers.
        Return nodes
    End Function
    
    Private Shared Sub VerifyParameterNodes(ByVal startNode As Node, ByVal endNode As Node)
        ' The order in which these checks are done is important.
        If startNode Is Nothing Then
            Throw New ArgumentException("Start node cannot be null")
        End If
        If endNode Is Nothing Then
            Throw New ArgumentException("End node cannot be null")
        End If

        If (Not startNode.Document.Equals(endNode.Document)) Then
            Throw New ArgumentException("Start node and end node must belong to the same document")
        End If

        If startNode.GetAncestor(NodeType.Body) Is Nothing OrElse endNode.GetAncestor(NodeType.Body) Is Nothing Then
            Throw New ArgumentException("Start node and end node must be a child or descendant of a body")
        End If

        ' Check the end node is after the start node in the DOM tree
        ' First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
        Dim startSection As Section = CType(startNode.GetAncestor(NodeType.Section), Section)
        Dim endSection As Section = CType(endNode.GetAncestor(NodeType.Section), Section)

        Dim startIndex As Integer = startSection.ParentNode.IndexOf(startSection)
        Dim endIndex As Integer = endSection.ParentNode.IndexOf(endSection)

        If startIndex = endIndex Then
            If startSection.Body.IndexOf(startNode) > endSection.Body.IndexOf(endNode) Then
                Throw New ArgumentException("The end node must be after the start node in the body")
            End If
        ElseIf startIndex > endIndex Then
            Throw New ArgumentException("The section of end node must be after the section start node")
        End If
    End Sub

    ''' <summary>
    ''' Checks if a node passed is an inline node.
    ''' </summary>
    Private Shared Function IsInline(ByVal node As Node) As Boolean
        ' Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
        Return ((node.GetAncestor(NodeType.Paragraph) IsNot Nothing OrElse node.GetAncestor(NodeType.Table) IsNot Nothing) AndAlso Not (node.NodeType = NodeType.Paragraph OrElse node.NodeType = NodeType.Table))
    End Function

    ''' <summary>
    ''' Removes the content before or after the marker in the cloned node depending on the type of marker.
    ''' </summary>
    Private Shared Sub ProcessMarker(ByVal cloneNode As CompositeNode, ByVal nodes As ArrayList, ByVal node As Node, ByVal isInclusive As Boolean, ByVal isStartMarker As Boolean, ByVal isEndMarker As Boolean)
        ' If we are dealing with a block level node just see if it should be included and add it to the list.
        If (Not IsInline(node)) Then
            ' Don't add the node twice if the markers are the same node
            If Not (isStartMarker AndAlso isEndMarker) Then
                If isInclusive Then
                    nodes.Add(cloneNode)
                End If
            End If
            Return
        End If

        ' If a marker is a FieldStart node check if it's to be included or not.
        ' We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        If node.NodeType = NodeType.FieldStart Then
            ' If the marker is a start node and is not be included then skip to the end of the field.
            ' If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
            If (isStartMarker AndAlso (Not isInclusive)) OrElse ((Not isStartMarker) AndAlso isInclusive) Then
                Do While node.NextSibling IsNot Nothing AndAlso node.NodeType <> NodeType.FieldEnd
                    node = node.NextSibling
                Loop

            End If
        End If

        ' If either marker is part of a comment then to include the comment itself we need to move the pointer forward to the Comment
        ' node found after the CommentRangeEnd node.
        If node.NodeType = NodeType.CommentRangeEnd Then
            Do While node.NextSibling IsNot Nothing AndAlso node.NodeType <> NodeType.Comment
                node = node.NextSibling
            Loop

        End If

        ' Find the corresponding node in our cloned node by index and return it.
        ' If the start and end node are the same some child nodes might already have been removed. Subtract the
        ' difference to get the right index.
        Dim indexDiff As Integer = node.ParentNode.ChildNodes.Count - cloneNode.ChildNodes.Count

        ' Child node count identical.
        If indexDiff = 0 Then
            node = cloneNode.ChildNodes(node.ParentNode.IndexOf(node))
        Else
            node = cloneNode.ChildNodes(node.ParentNode.IndexOf(node) - indexDiff)
        End If

        ' Remove the nodes up to/from the marker.
        Dim isSkip As Boolean = False
        Dim isProcessing As Boolean = True
        Dim isRemoving As Boolean = isStartMarker
        Dim nextNode As Node = cloneNode.FirstChild

        Do While isProcessing AndAlso nextNode IsNot Nothing
            Dim currentNode As Node = nextNode
            isSkip = False

            If currentNode.Equals(node) Then
                If isStartMarker Then
                    isProcessing = False
                    If isInclusive Then
                        isRemoving = False
                    End If
                Else
                    isRemoving = True
                    If isInclusive Then
                        isSkip = True
                    End If
                End If
            End If

            nextNode = nextNode.NextSibling
            If isRemoving AndAlso (Not isSkip) Then
                currentNode.Remove()
            End If
        Loop

        ' After processing the composite node may become empty. If it has don't include it.
        If Not (isStartMarker AndAlso isEndMarker) Then
            If cloneNode.HasChildNodes Then
                nodes.Add(cloneNode)
            End If
        End If

    End Sub
    
    Public Shared Function GenerateDocument(ByVal srcDoc As Document, ByVal nodes As ArrayList) As Document
        ' Create a blank document.
        Dim dstDoc As New Document()
        ' Remove the first paragraph from the empty document.
        dstDoc.FirstSection.Body.RemoveAllChildren()

        ' Import each node from the list into the new document. Keep the original formatting of the node.
        Dim importer As New NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting)

        For Each node As Node In nodes
            Dim importNode As Node = importer.ImportNode(node, True)
            dstDoc.FirstSection.Body.AppendChild(importNode)
        Next node

        ' Return the generated document.
        Return dstDoc
    End Function
    
    Public Shared Function ParagraphsByStyleName(ByVal doc As Document, ByVal styleName As String) As ArrayList
        ' Create an array to collect paragraphs of the specified style.
        Dim paragraphsWithStyle As New ArrayList()
        ' Get all paragraphs from the document.
        Dim paragraphs As NodeCollection = doc.GetChildNodes(NodeType.Paragraph, True)
        ' Look through all paragraphs to find those with the specified style.
        For Each paragraph As Paragraph In paragraphs
            If paragraph.ParagraphFormat.Style.Name = styleName Then
                paragraphsWithStyle.Add(paragraph)
            End If
        Next paragraph
        Return paragraphsWithStyle
    End Function
End Class
