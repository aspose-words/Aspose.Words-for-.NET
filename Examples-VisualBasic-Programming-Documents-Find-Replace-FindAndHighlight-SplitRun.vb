' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
