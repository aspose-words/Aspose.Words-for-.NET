// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
string fileName = "TestFile.Destination.doc";

Document dstDoc = new Document(dataDir + fileName);
Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

// Set the appended document to start on a new page.
srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;

// Append the source document using the original styles found in the source document.
dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
dstDoc.Save(dataDir);
