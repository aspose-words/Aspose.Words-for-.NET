// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
/// <summary>
/// Represents a facade object for a merge field in a Microsoft Word document.
/// </summary>
internal class MergeField
{
    internal MergeField(FieldStart fieldStart)
    {
        if (fieldStart.Equals(null))
            throw new ArgumentNullException("fieldStart");
        if (!fieldStart.FieldType.Equals(FieldType.FieldMergeField))
            throw new ArgumentException("Field start type must be FieldMergeField.");

        mFieldStart = fieldStart;

        // Find the field separator node.
        mFieldSeparator = fieldStart.GetField().Separator;
        if (mFieldSeparator == null)
            throw new InvalidOperationException("Cannot find field separator.");

        mFieldEnd = fieldStart.GetField().End;
    }

    /// <summary>
    /// Gets or sets the name of the merge field.
    /// </summary>
    internal string Name
    {
        get
        {
            return ((FieldStart)mFieldStart).GetField().Result.Replace("«", "").Replace("»", "");
        }
        set
        {
            // Merge field name is stored in the field result which is a Run
            // node between field separator and field end.
            Run fieldResult = (Run)mFieldSeparator.NextSibling;
            fieldResult.Text = string.Format("«{0}»", value);

            // But sometimes the field result can consist of more than one run, delete these runs.
            RemoveSameParent(fieldResult.NextSibling, mFieldEnd);

            UpdateFieldCode(value);
        }
    }

    private void UpdateFieldCode(string fieldName)
    {
        // Field code is stored in a Run node between field start and field separator.
        Run fieldCode = (Run)mFieldStart.NextSibling;

        Match match = gRegex.Match(((FieldStart)mFieldStart).GetField().GetFieldCode());

        string newFieldCode = string.Format(" {0}{1} ", match.Groups["start"].Value, fieldName);
        fieldCode.Text = newFieldCode;

        // But sometimes the field code can consist of more than one run, delete these runs.
        RemoveSameParent(fieldCode.NextSibling, mFieldSeparator);
    }

    /// <summary>
    /// Removes nodes from start up to but not including the end node.
    /// Start and end are assumed to have the same parent.
    /// </summary>
    private static void RemoveSameParent(Node startNode, Node endNode)
    {
        if ((endNode != null) && ((Aspose.Words.Node)startNode.ParentNode != (Aspose.Words.Node)endNode.ParentNode))
            throw new ArgumentException("Start and end nodes are expected to have the same parent.");

        Node curChild = startNode;
        while ((curChild != null) && (curChild != endNode))
        {
            Node nextChild = curChild.NextSibling;
            curChild.Remove();
            curChild = nextChild;
        }
    }

    private readonly Node mFieldStart;
    private readonly Node mFieldSeparator;
    private readonly Node mFieldEnd;

    private static readonly Regex gRegex = new Regex(@"\s*(?<start>MERGEFIELD\s|)(\s|)(?<name>\S+)\s+");
}
