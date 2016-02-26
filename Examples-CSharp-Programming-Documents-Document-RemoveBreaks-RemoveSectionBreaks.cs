// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
private static void RemoveSectionBreaks(Document doc)
{
    // Loop through all sections starting from the section that precedes the last one 
    // and moving to the first section.
    for (int i = doc.Sections.Count - 2; i >= 0; i--)
    {
        // Copy the content of the current section to the beginning of the last section.
        doc.LastSection.PrependContent(doc.Sections[i]);
        // Remove the copied section.
        doc.Sections[i].Remove();
    }
}
