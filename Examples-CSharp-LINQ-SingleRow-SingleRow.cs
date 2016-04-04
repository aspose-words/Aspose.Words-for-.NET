// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LINQ(); 
string fileName = "SingleRow.doc";
// Load the template document.
Document doc = new Document(dataDir + fileName);

// Load the photo and read all bytes.
byte[] imgdata = System.IO.File.ReadAllBytes(dataDir + "photo.png");
           
// Create a Reporting Engine.
ReportingEngine engine = new ReportingEngine();
            
// Execute the build report.
engine.BuildReport(doc, Common.GetManager(), "manager");

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);

// Save the finished document to disk.
doc.Save(dataDir);
