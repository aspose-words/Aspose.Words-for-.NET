'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Text.RegularExpressions
Imports System.Collections
Imports System.Drawing
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Public Class FindAndHighlight
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_FindAndReplace()

        Dim doc As New Document(dataDir & "TestFile.doc")

        ' We want the "your document" phrase to be highlighted.
        Dim regex As New Regex("your document", RegexOptions.IgnoreCase)
        doc.Range.Replace(regex, New ReplaceEvaluatorFindAndHighlight(), False)

        ' Save the output document.
        doc.Save(dataDir & "TestFile Out.doc")

        Console.WriteLine(vbNewLine & "Text highlighted successfully." & vbNewLine & "File saved at " + dataDir + "TestFile Out.doc")
    End Sub

    Private Class ReplaceEvaluatorFindAndHighlight
        Implements IReplacingCallback
        ''' <summary>
        ''' This method is called by the Aspose.Words find and replace engine for each match.
        ''' This method highlights the match string, even if it spans multiple runs.
        ''' </summary>
        Private Function IReplacingCallback_Replacing(ByVal e As ReplacingArgs) As ReplaceAction Implements IReplacingCallback.Replacing
            ' This is a Run node that contains either the beginning or the complete match.
            Dim currentNode As Node = e.MatchNode

            ' The first (and may be the only) run can contain text before the match, 
            ' in this case it is necessary to split the run.
            If e.MatchOffset > 0 Then
                currentNode = SplitRun(CType(currentNode, Run), e.MatchOffset)
            End If

            ' This array is used to store all nodes of the match for further highlighting.
            Dim runs As New ArrayList()

            ' Find all runs that contain parts of the match string.
            Dim remainingLength As Integer = e.Match.Value.Length
            Do While (remainingLength > 0) AndAlso (currentNode IsNot Nothing) AndAlso (currentNode.GetText().Length <= remainingLength)
                runs.Add(currentNode)
                remainingLength = remainingLength - currentNode.GetText().Length

                ' Select the next Run node. 
                ' Have to loop because there could be other nodes such as BookmarkStart etc.
                Do
                    currentNode = currentNode.NextSibling
                Loop While (currentNode IsNot Nothing) AndAlso (currentNode.NodeType <> NodeType.Run)
            Loop

            ' Split the last run that contains the match if there is any text left.
            If (currentNode IsNot Nothing) AndAlso (remainingLength > 0) Then
                SplitRun(CType(currentNode, Run), remainingLength)
                runs.Add(currentNode)
            End If

            ' Now highlight all runs in the sequence.
            For Each run As Run In runs
                run.Font.HighlightColor = Color.Yellow
            Next run

            ' Signal to the replace engine to do nothing because we have already done all what we wanted.
            Return ReplaceAction.Skip
        End Function
    End Class

    ''' <summary>
    ''' Splits text of the specified run into two runs.
    ''' Inserts the new run just after the specified run.
    ''' </summary>
    Private Shared Function SplitRun(ByVal run As Run, ByVal position As Integer) As Run
        Dim afterRun As Run = CType(run.Clone(True), Run)
        afterRun.Text = run.Text.Substring(position)
        run.Text = run.Text.Substring(0, position)
        run.ParentNode.InsertAfter(afterRun, run)
        Return afterRun
    End Function
End Class
