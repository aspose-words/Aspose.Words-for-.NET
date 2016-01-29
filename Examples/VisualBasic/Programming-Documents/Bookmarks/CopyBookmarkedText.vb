Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Public Class CopyBookmarkedText
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithBookmarks()
        Dim fileName As String = "Template.doc"

        ' Load the source document.
        Dim srcDoc As New Document(dataDir & fileName)

        ' This is the bookmark whose content we want to copy.
        Dim srcBookmark As Bookmark = srcDoc.Range.Bookmarks("ntf010145060")

        ' We will be adding to this document.
        Dim dstDoc As New Document()

        ' Let's say we will be appending to the end of the body of the last section.
        Dim dstNode As CompositeNode = dstDoc.LastSection.Body

        ' It is a good idea to use this import context object because multiple nodes are being imported.
        ' If you import multiple times without a single context, it will result in many styles created.
        Dim importer As New NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting)

        ' Do it once.
        AppendBookmarkedText(importer, srcBookmark, dstNode)

        ' Do it one more time for fun.
        AppendBookmarkedText(importer, srcBookmark, dstNode)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the finished document.
        dstDoc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Bookmark copied successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

    Private Shared Sub AppendBookmarkedText(ByVal importer As NodeImporter, ByVal srcBookmark As Bookmark, ByVal dstNode As CompositeNode)
        ' This is the paragraph that contains the beginning of the bookmark.
        Dim startPara As Paragraph = CType(srcBookmark.BookmarkStart.ParentNode, Paragraph)

        ' This is the paragraph that contains the end of the bookmark.
        Dim endPara As Paragraph = CType(srcBookmark.BookmarkEnd.ParentNode, Paragraph)

        If (startPara Is Nothing) OrElse (endPara Is Nothing) Then
            Throw New InvalidOperationException("Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.")
        End If

        ' Limit ourselves to a reasonably simple scenario.
        If startPara.ParentNode IsNot endPara.ParentNode Then
            Throw New InvalidOperationException("Start and end paragraphs have different parents, cannot handle this scenario yet.")
        End If

        ' We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
        ' therefore the node at which we stop is one after the end paragraph.
        Dim endNode As Node = endPara.NextSibling

        ' This is the loop to go through all paragraph-level nodes in the bookmark.
        Dim curNode As Node = startPara
        Do While curNode IsNot endNode
            ' This creates a copy of the current node and imports it (makes it valid) in the context
            ' of the destination document. Importing means adjusting styles and list identifiers correctly.
            Dim newNode As Node = importer.ImportNode(curNode, True)

            ' Now we simply append the new node to the destination.
            dstNode.AppendChild(newNode)
            curNode = curNode.NextSibling
        Loop
    End Sub
End Class
