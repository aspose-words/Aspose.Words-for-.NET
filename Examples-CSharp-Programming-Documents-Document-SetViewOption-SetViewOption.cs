// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
// Load the template document.
Document doc = new Document(dataDir + "TestFile.doc");
// Set view option.
doc.ViewOptions.ViewType = ViewType.PageLayout;
doc.ViewOptions.ZoomPercent = 50;

dataDir = dataDir + "TestFile.SetZoom_out_.doc";
// Save the finished document.
doc.Save(dataDir);
