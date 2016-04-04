// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
// Load the template document.
Document doc = new Document(dataDir + "TestFile.doc");
// Get styles collection from document.
StyleCollection styles = doc.Styles;
string styleName = "";
// Iterate through all the styles.
foreach (Style style in styles)
{
    if (styleName == "")
    {
        styleName = style.Name;
    }
    else
    {
        styleName = styleName + ", " + style.Name;
    }
}
