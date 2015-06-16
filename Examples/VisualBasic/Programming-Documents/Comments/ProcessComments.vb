'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Public Class ProcessComments
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithComments()

        ' Open the document.
        Dim doc As New Document(dataDir & "TestFile.doc")

        For Each comment As String In ExtractComments(doc)
            Console.Write(comment)
        Next comment

        ' Remove comments by the "pm" author.
        RemoveComments(doc, "pm")
        Console.WriteLine("Comments from ""pm"" are removed!")

        ' Extract the information about the comments of the "ks" author.
        For Each comment As String In ExtractComments(doc, "ks")
            Console.Write(comment)
        Next comment

        ' Remove all comments.
        RemoveComments(doc)
        Console.WriteLine("All comments are removed!")

        ' Save the document.
        doc.Save(dataDir & "Test File Out.doc")

        Console.WriteLine(vbNewLine & "Comments extracted and removed successfully." & vbNewLine & "File saved at " + dataDir + "Test File Out.doc")
    End Sub

    Private Shared Function ExtractComments(ByVal doc As Document) As ArrayList
        Dim collectedComments As New ArrayList()
        ' Collect all comments in the document
        Dim comments As NodeCollection = doc.GetChildNodes(NodeType.Comment, True)
        ' Look through all comments and gather information about them.
        For Each comment As Comment In comments
            collectedComments.Add(comment.Author & " " & comment.DateTime & " " & comment.ToString(SaveFormat.Text))
        Next comment
        Return collectedComments
    End Function

    Private Shared Function ExtractComments(ByVal doc As Document, ByVal authorName As String) As ArrayList
        Dim collectedComments As New ArrayList()
        ' Collect all comments in the document
        Dim comments As NodeCollection = doc.GetChildNodes(NodeType.Comment, True)
        ' Look through all comments and gather information about those written by the authorName author.
        For Each comment As Comment In comments
            If comment.Author = authorName Then
                collectedComments.Add(comment.Author & " " & comment.DateTime & " " & comment.ToString(SaveFormat.Text))
            End If
        Next comment
        Return collectedComments
    End Function

    Private Shared Sub RemoveComments(ByVal doc As Document)
        ' Collect all comments in the document
        Dim comments As NodeCollection = doc.GetChildNodes(NodeType.Comment, True)
        ' Remove all comments.
        comments.Clear()
    End Sub

    Private Shared Sub RemoveComments(ByVal doc As Document, ByVal authorName As String)
        ' Collect all comments in the document
        Dim comments As NodeCollection = doc.GetChildNodes(NodeType.Comment, True)
        ' Look through all comments and remove those written by the authorName author.
        For i As Integer = comments.Count - 1 To 0 Step -1
            Dim comment As Comment = CType(comments(i), Comment)
            If comment.Author = authorName Then
                comment.Remove()
            End If
        Next i
    End Sub
End Class
