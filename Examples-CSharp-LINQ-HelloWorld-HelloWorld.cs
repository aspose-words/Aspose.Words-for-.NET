// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LINQ();

string fileName = "HelloWorld.doc";
// Load the template document.
Document doc = new Document(dataDir + fileName);

// Create an instance of sender class to set it's properties.
Sender sender = new Sender { Name = "LINQ Reporting Engine", Message = "Hello World" };
            
// Create a Reporting Engine.
ReportingEngine engine = new ReportingEngine();
            
// Execute the build report.
engine.BuildReport(doc, sender, "sender");

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

// Save the finished document to disk.
doc.Save(dataDir);
