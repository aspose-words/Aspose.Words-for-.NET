// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
/// <summary>
/// Clones and copies headers/footers form the previous section to the specified section.
/// </summary>
private static void CopyHeadersFootersFromPreviousSection(Section section)
{
    Section previousSection = (Section)section.PreviousSibling;

    if (previousSection == null)
        return;

    section.HeadersFooters.Clear();

    foreach (HeaderFooter headerFooter in previousSection.HeadersFooters)
        section.HeadersFooters.Add(headerFooter.Clone(true));
}
