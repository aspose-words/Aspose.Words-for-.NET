// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LINQ(); 
string fileName = "InTableWithFilteringGroupingSorting.doc";
// Load the template document.
Document doc = new Document(dataDir + fileName);

// Create a Reporting Engine.
ReportingEngine engine = new ReportingEngine();
            
// Execute the build report.
engine.BuildReport(doc, Common.GetContracts(), "contracts");

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

// Save the finished document to disk.
doc.Save(dataDir);
