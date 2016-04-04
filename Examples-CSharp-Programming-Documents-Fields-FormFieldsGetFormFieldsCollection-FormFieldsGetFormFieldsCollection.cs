// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithFields();

Document doc = new Document(dataDir + "FormFields.doc");
FormFieldCollection formFields = doc.Range.FormFields;

