// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 

// Create the Dataset and read the XML.
DataSet customersDs = new DataSet();
customersDs.ReadXml(dataDir + "Customers.xml");

string fileName = "TestFile XML.doc";
// Open a template document.
Document doc = new Document(dataDir + fileName);

// Execute mail merge to fill the template with data from XML using DataTable.
doc.MailMerge.Execute(customersDs.Tables["Customer"]);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the output document.
doc.Save(dataDir);
