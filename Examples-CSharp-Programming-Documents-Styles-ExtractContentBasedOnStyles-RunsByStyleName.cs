// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
public static ArrayList RunsByStyleName(Document doc, string styleName)
{
    // Create an array to collect runs of the specified style.
    ArrayList runsWithStyle = new ArrayList();
    // Get all runs from the document.
    NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
    // Look through all runs to find those with the specified style.
    foreach (Run run in runs)
    {
        if (run.Font.Style.Name == styleName)
            runsWithStyle.Add(run);
    }
    return runsWithStyle;
}
