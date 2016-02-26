// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

string fileName = "TestFile.doc";
            
// Open the document.
Document doc = new Document(dataDir + fileName);

// Remove the page and section breaks from the document.
// In Aspose.Words section breaks are represented as separate Section nodes in the document.
// To remove these separate sections the sections are combined.
RemovePageBreaks(doc);
RemoveSectionBreaks(doc);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the document.
doc.Save(dataDir);
