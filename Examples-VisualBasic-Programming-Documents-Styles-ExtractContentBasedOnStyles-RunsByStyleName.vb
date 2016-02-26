' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Public Shared Function RunsByStyleName(ByVal doc As Document, ByVal styleName As String) As ArrayList
    ' Create an array to collect runs of the specified style.
    Dim runsWithStyle As New ArrayList()
    ' Get all runs from the document.
    Dim runs As NodeCollection = doc.GetChildNodes(NodeType.Run, True)
    ' Look through all runs to find those with the specified style.
    For Each run As Run In runs
        If run.Font.Style.Name = styleName Then
            runsWithStyle.Add(run)
        End If
    Next run
    Return runsWithStyle
End Function
