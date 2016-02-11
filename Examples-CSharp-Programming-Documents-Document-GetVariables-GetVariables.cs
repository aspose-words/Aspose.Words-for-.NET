// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
// Load the template document.
Document doc = new Document(dataDir + "TestFile.doc");
string variables = "";
foreach (DictionaryEntry entry in doc.Variables)
{
    string name = entry.Key.ToString();
    string value = entry.Value.ToString();
    if (variables == "")
    {
        // Do something useful.
        variables = "Name: " + name + "," + "Value: {1}" + value;
    }
    else
    {
        variables = variables + "Name: " + name + "," + "Value: {1}" + value;
    }
}
