// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
DataSet ds = new DataSet();

ds.ReadXml(dataDir + "Vendors.xml");

// Open a template document.
Document doc = new Document(dataDir + "VendorTemplate.doc");

doc.MailMerge.UseNonMergeFields = true;

// Execute mail merge to fill the template with data from XML using DataSet.
doc.MailMerge.ExecuteWithRegions(ds);
dataDir = dataDir + "MailMergeUsingMustacheSyntax_out_.docx";
// Save the output document.
doc.Save(dataDir);
