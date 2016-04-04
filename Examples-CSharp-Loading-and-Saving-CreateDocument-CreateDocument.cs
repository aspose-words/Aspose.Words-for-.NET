// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

// Initialize a Document.
Document doc = new Document();
            
// Use a document builder to add content to the document.
DocumentBuilder builder = new DocumentBuilder(doc);
builder.Writeln("Hello World!");  
        
dataDir  = dataDir + "CreateDocument_out_.docx";
// Save the document to disk.
doc.Save(dataDir);
           
