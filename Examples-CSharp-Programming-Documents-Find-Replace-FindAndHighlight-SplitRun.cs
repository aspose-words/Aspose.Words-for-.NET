// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
/// <summary>
/// Splits text of the specified run into two runs.
/// Inserts the new run just after the specified run.
/// </summary>
private static Run SplitRun(Run run, int position)
{
    Run afterRun = (Run)run.Clone(true);
    afterRun.Text = run.Text.Substring(position);
    run.Text = run.Text.Substring(0, position);
    run.ParentNode.InsertAfter(afterRun, run);
    return afterRun;
}
