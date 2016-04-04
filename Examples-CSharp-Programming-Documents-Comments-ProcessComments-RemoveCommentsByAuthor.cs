// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
static void RemoveComments(Document doc, string authorName)
{
    // Collect all comments in the document
    NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
    // Look through all comments and remove those written by the authorName author.
    for (int i = comments.Count - 1; i >= 0; i--)
    {
        Comment comment = (Comment)comments[i];
        if (comment.Author == authorName)
            comment.Remove();
    }
}
