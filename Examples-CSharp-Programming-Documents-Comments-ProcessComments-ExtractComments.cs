// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
static ArrayList ExtractComments(Document doc)
{
    ArrayList collectedComments = new ArrayList();
    // Collect all comments in the document
    NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
    // Look through all comments and gather information about them.
    foreach (Comment comment in comments)
    {
        collectedComments.Add(comment.Author + " " + comment.DateTime + " " + comment.ToString(SaveFormat.Text));
    }
    return collectedComments;
}
