// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_FindAndReplace();
string fileName = "TestFile.doc";

Document doc = new Document(dataDir + fileName);

// We want the "your document" phrase to be highlighted.
Regex regex = new Regex("your document", RegexOptions.IgnoreCase);
doc.Range.Replace(regex, new ReplaceEvaluatorFindAndHighlight(), false);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the output document.
doc.Save(dataDir);
