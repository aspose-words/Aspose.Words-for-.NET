// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithFields();

Document doc = new Document(dataDir + "FormFields.doc");
FormField formField = doc.Range.FormFields[3];

if (formField.Type.Equals(FieldType.FieldFormTextInput))
    formField.Result = "My name is " + formField.Name;
