// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithFields();

string fileName = "TestFile.doc";
Document doc = new Document(dataDir + fileName);

// Pass the appropriate parameters to convert PAGE fields encountered to static text only in the body of the first section.
FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body, FieldType.FieldPage);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the document with fields transformed to disk.
doc.Save(dataDir);
