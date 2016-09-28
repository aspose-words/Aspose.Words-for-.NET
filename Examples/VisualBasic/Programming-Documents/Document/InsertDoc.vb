Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.MailMerging
Imports Aspose.Words.Replacing
Imports System.Text.RegularExpressions

Class InsertDoc
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Invokes the InsertDocument method shown above to insert a document at a bookmark.
        InsertDocumentAtBookmark(dataDir)
        InsertDocumentAtMailMerge(dataDir)
        InsertDocumentAtReplace(dataDir)
    End Sub
    Public Shared Sub InsertDocumentAtReplace(dataDir As String)
        ' ExStart:InsertDocumentAtReplace
        Dim mainDoc As New Document(dataDir & Convert.ToString("InsertDocument1.doc"))
        Dim options As New FindReplaceOptions()
        options.ReplacingCallback = New InsertDocumentAtReplaceHandler()
        mainDoc.Range.Replace(New Regex("\[MY_DOCUMENT\]"), "", options)
        dataDir = dataDir & Convert.ToString("InsertDocumentAtReplace_out_.doc")
        mainDoc.Save(dataDir)
        ' ExEnd:InsertDocumentAtReplace
        Console.WriteLine(Convert.ToString(vbLf & "Document inserted successfully at a replace." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertDocumentAtBookmark(dataDir As String)
        ' ExStart:InsertDocumentAtBookmark         
        Dim mainDoc As New Document(dataDir & Convert.ToString("InsertDocument1.doc"))
        Dim subDoc As New Document(dataDir & Convert.ToString("InsertDocument2.doc"))

        Dim bookmark As Bookmark = mainDoc.Range.Bookmarks("insertionPlace")
        InsertDocument(bookmark.BookmarkStart.ParentNode, subDoc)
        dataDir = dataDir & Convert.ToString("InsertDocumentAtBookmark_out_.doc")
        mainDoc.Save(dataDir)
        ' ExEnd:InsertDocumentAtBookmark
        Console.WriteLine(Convert.ToString(vbLf & "Document inserted successfully at a bookmark." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertDocumentAtMailMerge(dataDir As String)
        ' ExStart:InsertDocumentAtMailMerge   
        ' Open the main document.
        Dim mainDoc As New Document(dataDir & Convert.ToString("InsertDocument1.doc"))

        ' Add a handler to MergeField event
        mainDoc.MailMerge.FieldMergingCallback = New InsertDocumentAtMailMergeHandler()

        ' The main document has a merge field in it called "Document_1".
        ' The corresponding data for this field contains fully qualified path to the document
        ' that should be inserted to this field.
        mainDoc.MailMerge.Execute(New String() {"Document_1"}, New String() {dataDir & Convert.ToString("InsertDocument2.doc")})
        dataDir = dataDir & Convert.ToString("InsertDocumentAtMailMerge_out_.doc")
        mainDoc.Save(dataDir)
        ' ExEnd:InsertDocumentAtMailMerge 
        Console.WriteLine(Convert.ToString(vbLf & "Document inserted successfully at mail merge." & vbLf & "File saved at ") & dataDir)
    End Sub
    ' ExStart:InsertDocument
    ''' <summary>
    ''' Inserts content of the external document after the specified node.
    ''' Section breaks and section formatting of the inserted document are ignored.
    ''' </summary>
    ''' <param name="insertAfterNode">Node in the destination document after which the content
    ''' should be inserted. This node should be a block level node (paragraph or table).</param>
    ''' <param name="srcDoc">The document to insert.</param>
    Private Shared Sub InsertDocument(insertAfterNode As Node, srcDoc As Document)
        ' Make sure that the node is either a paragraph or table.
        If (Not insertAfterNode.NodeType.Equals(NodeType.Paragraph)) And (Not insertAfterNode.NodeType.Equals(NodeType.Table)) Then
            Throw New ArgumentException("The destination node should be either a paragraph or table.")
        End If

        ' We will be inserting into the parent of the destination paragraph.
        Dim dstStory As CompositeNode = insertAfterNode.ParentNode

        ' This object will be translating styles and lists during the import.
        Dim importer As New NodeImporter(srcDoc, insertAfterNode.Document, ImportFormatMode.KeepSourceFormatting)

        ' Loop through all sections in the source document.
        For Each srcSection As Section In srcDoc.Sections
            ' Loop through all block level nodes (paragraphs and tables) in the body of the section.
            For Each srcNode As Node In srcSection.Body
                ' Let's skip the node if it is a last empty paragraph in a section.
                If srcNode.NodeType.Equals(NodeType.Paragraph) Then
                    Dim para As Paragraph = DirectCast(srcNode, Paragraph)
                    If para.IsEndOfSection AndAlso Not para.HasChildNodes Then
                        Continue For
                    End If
                End If

                ' This creates a clone of the node, suitable for insertion into the destination document.
                Dim newNode As Node = importer.ImportNode(srcNode, True)

                ' Insert new node after the reference node.
                dstStory.InsertAfter(newNode, insertAfterNode)
                insertAfterNode = newNode
            Next
        Next
    End Sub
    ' ExEnd:InsertDocument
    ' ExStart:InsertDocumentWithSectionFormatting
    ''' <summary>
    ''' Inserts content of the external document after the specified node.
    ''' </summary>
    ''' <param name="insertAfterNode">Node in the destination document after which the content
    ''' should be inserted. This node should be a block level node (paragraph or table).</param>
    ''' <param name="srcDoc">The document to insert.</param>
    Private Shared Sub InsertDocumentWithSectionFormatting(insertAfterNode As Node, srcDoc As Document)
        ' Make sure that the node is either a pargraph or table.
        If (Not insertAfterNode.NodeType.Equals(NodeType.Paragraph)) And (Not insertAfterNode.NodeType.Equals(NodeType.Table)) Then
            Throw New ArgumentException("The destination node should be either a paragraph or table.")
        End If

        ' Document to insert srcDoc into.
        Dim dstDoc As Document = DirectCast(insertAfterNode.Document, Document)
        ' To retain section formatting, split the current section into two at the marker node and then import the content from srcDoc as whole sections.
        ' The section of the node which the insert marker node belongs to
        Dim currentSection As Section = DirectCast(insertAfterNode.GetAncestor(NodeType.Section), Section)

        ' Don't clone the content inside the section, we just want the properties of the section retained.
        Dim cloneSection As Section = DirectCast(currentSection.Clone(False), Section)

        ' However make sure the clone section has a body, but no empty first paragraph.
        cloneSection.EnsureMinimum()
        cloneSection.Body.FirstParagraph.Remove()

        ' Insert the cloned section into the document after the original section.
        insertAfterNode.Document.InsertAfter(cloneSection, currentSection)

        ' Append all nodes after the marker node to the new section. This will split the content at the section level at
        ' the marker so the sections from the other document can be inserted directly.
        Dim currentNode As Node = insertAfterNode.NextSibling
        While currentNode IsNot Nothing
            Dim nextNode As Node = currentNode.NextSibling
            cloneSection.Body.AppendChild(currentNode)
            currentNode = nextNode
        End While

        ' This object will be translating styles and lists during the import.
        Dim importer As New NodeImporter(srcDoc, dstDoc, ImportFormatMode.UseDestinationStyles)

        ' Loop through all sections in the source document.
        For Each srcSection As Section In srcDoc.Sections
            Dim newNode As Node = importer.ImportNode(srcSection, True)

            ' Append each section to the destination document. Start by inserting it after the split section.
            dstDoc.InsertAfter(newNode, currentSection)
            currentSection = DirectCast(newNode, Section)
        Next
    End Sub
    ' ExEnd:InsertDocumentWithSectionFormatting
    ' ExStart:InsertDocumentAtMailMergeHandler
    Private Class InsertDocumentAtMailMergeHandler
        Implements IFieldMergingCallback
        ''' <summary>
        ''' This handler makes special processing for the "Document_1" field.
        ''' The field value contains the path to load the document. 
        ''' We load the document and insert it into the current merge field.
        ''' </summary>
        Private Sub IFieldMergingCallback_FieldMerging(e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
            If e.DocumentFieldName = "Document_1" Then
                ' Use document builder to navigate to the merge field with the specified name.
                Dim builder As New DocumentBuilder(e.Document)
                builder.MoveToMergeField(e.DocumentFieldName)

                ' The name of the document to load and insert is stored in the field value.
                Dim subDoc As New Aspose.Words.Document(DirectCast(e.FieldValue, String))

                ' Insert the document.
                InsertDocument(builder.CurrentParagraph, subDoc)

                ' The paragraph that contained the merge field might be empty now and you probably want to delete it.
                If Not builder.CurrentParagraph.HasChildNodes Then
                    builder.CurrentParagraph.Remove()
                End If

                ' Indicate to the mail merge engine that we have inserted what we wanted.
                e.Text = Nothing
            End If
        End Sub

        Private Sub IFieldMergingCallback_ImageFieldMerging(args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
            ' Do nothing.
        End Sub
    End Class
    ' ExEnd:InsertDocumentAtMailMergeHandler
    ' ExStart:InsertDocumentAtMailMergeBlobHandler
    Private Class InsertDocumentAtMailMergeBlobHandler
        Implements IFieldMergingCallback
        ''' <summary>
        ''' This handler makes special processing for the "Document_1" field.
        ''' The field value contains the path to load the document.
        ''' We load the document and insert it into the current merge field.
        ''' </summary>
        Private Sub IFieldMergingCallback_FieldMerging(e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
            If e.DocumentFieldName = "Document_1" Then
                ' Use document builder to navigate to the merge field with the specified name.
                Dim builder As New DocumentBuilder(e.Document)
                builder.MoveToMergeField(e.DocumentFieldName)

                ' Load the document from the blob field.
                Dim stream As New MemoryStream(DirectCast(e.FieldValue, Byte()))
                Dim subDoc As New Document(stream)

                ' Insert the document.
                InsertDocument(builder.CurrentParagraph, subDoc)

                ' The paragraph that contained the merge field might be empty now and you probably want to delete it.
                If Not builder.CurrentParagraph.HasChildNodes Then
                    builder.CurrentParagraph.Remove()
                End If

                ' Indicate to the mail merge engine that we have inserted what we wanted.
                e.Text = Nothing
            End If
        End Sub

        Private Sub IFieldMergingCallback_ImageFieldMerging(args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
            ' Do nothing.
        End Sub
    End Class
    ' ExEnd:InsertDocumentAtMailMergeBlobHandler
    ' ExStart:InsertDocumentAtReplaceHandler
    Private Class InsertDocumentAtReplaceHandler
        Implements IReplacingCallback
        Private Function IReplacingCallback_Replacing(e As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
            Dim subDoc As New Document(RunExamples.GetDataDir_WorkingWithDocument() + "InsertDocument2.doc")

            ' Insert a document after the paragraph, containing the match text.
            Dim para As Paragraph = DirectCast(e.MatchNode.ParentNode, Paragraph)
            InsertDocument(para, subDoc)

            ' Remove the paragraph with the match text.
            para.Remove()

            Return ReplaceAction.Skip
        End Function
    End Class
    ' ExEnd:InsertDocumentAtReplaceHandler

End Class

