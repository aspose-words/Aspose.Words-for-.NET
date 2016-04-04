// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
Document doc = new Document(dataDir + "MailMerge.AlternatingRows.doc");

// Add a handler for the MergeField event.
doc.MailMerge.FieldMergingCallback = new HandleMergeFieldAlternatingRows();

// Execute mail merge with regions.
DataTable dataTable = GetSuppliersDataTable();
doc.MailMerge.ExecuteWithRegions(dataTable);
dataDir = dataDir + "MailMerge.AlternatingRows_out_.doc";
doc.Save(dataDir);
