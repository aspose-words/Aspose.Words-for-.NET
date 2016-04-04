// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
static ArrayList ExtractComments(Document doc, string authorName)
{
    ArrayList collectedComments = new ArrayList();
    // Collect all comments in the document
    NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
    // Look through all comments and gather information about those written by the authorName author.
    foreach (Comment comment in comments)
    {
        if (comment.Author == authorName)
            collectedComments.Add(comment.Author + " " + comment.DateTime + " " + comment.ToString(SaveFormat.Text));
    }
    return collectedComments;
}
