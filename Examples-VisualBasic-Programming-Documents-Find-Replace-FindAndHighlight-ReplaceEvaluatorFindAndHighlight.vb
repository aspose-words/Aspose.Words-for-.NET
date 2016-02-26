' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
