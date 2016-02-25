// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithHyperlink();
string NewUrl = @"http://www.aspose.com";
string NewName = "Aspose - The .NET & Java Component Publisher";
Document doc = new Document(dataDir + "ReplaceHyperlinks.doc");

// Hyperlinks in a Word documents are fields, select all field start nodes so we can find the hyperlinks.
NodeList fieldStarts = doc.SelectNodes("//FieldStart");
foreach (FieldStart fieldStart in fieldStarts)
{
    if (fieldStart.FieldType.Equals(FieldType.FieldHyperlink))
    {
        // The field is a hyperlink field, use the "facade" class to help to deal with the field.
        Hyperlink hyperlink = new Hyperlink(fieldStart);

        // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
        if (hyperlink.IsLocal)
            continue;

        // The Hyperlink class allows to set the target URL and the display name
        // of the link easily by setting the properties.
        hyperlink.Target = NewUrl;
        hyperlink.Name = NewName;
    }
}
dataDir = dataDir + "ReplaceHyperlinks_out_.doc";
doc.Save(dataDir);
