// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
static void RemoveComments(Document doc)
{
    // Collect all comments in the document
    NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
    // Remove all comments.
    comments.Clear();
}
