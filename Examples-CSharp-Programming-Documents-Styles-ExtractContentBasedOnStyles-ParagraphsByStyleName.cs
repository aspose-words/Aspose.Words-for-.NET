// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
public static ArrayList ParagraphsByStyleName(Document doc, string styleName)
{
    // Create an array to collect paragraphs of the specified style.
    ArrayList paragraphsWithStyle = new ArrayList();
    // Get all paragraphs from the document.
    NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
    // Look through all paragraphs to find those with the specified style.
    foreach (Paragraph paragraph in paragraphs)
    {
        if (paragraph.ParagraphFormat.Style.Name == styleName)
            paragraphsWithStyle.Add(paragraph);
    }
    return paragraphsWithStyle;
}
