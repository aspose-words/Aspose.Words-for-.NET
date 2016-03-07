// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 

string fileName = "TestFile Empty.doc";
// Open the document.
Document doc = new Document(dataDir + fileName);

// Create a dummy data source containing no data.
DataSet data = new DataSet();
// Set the appropriate mail merge clean up options to remove any unused regions from the document.
doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions;
//doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields;
//doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveStaticFields;
//doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveEmptyParagraphs;           
//doc.MailMerge.CleanupOptions |= MailMergeCleanupOptions.RemoveUnusedFields;
// Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
// automatically as they are unused.
doc.MailMerge.ExecuteWithRegions(data);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the output document to disk.
doc.Save(dataDir);
