// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            
// The encoding of the text file is automatically detected.
Document doc = new Document(dataDir + "LoadTxt.txt");

// Save as any Aspose.Words supported format, such as DOCX.  
dataDir = dataDir + "LoadTxt_out_.docx";
doc.Save(dataDir);
