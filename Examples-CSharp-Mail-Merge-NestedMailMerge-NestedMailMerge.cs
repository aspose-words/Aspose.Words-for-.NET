// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
            
// Create the Dataset and read the XML.
DataSet pizzaDs = new DataSet();

// Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
// To see examples of how to set up relations manually check the corresponding documentation of this sample
pizzaDs.ReadXml(dataDir + "CustomerData.xml");

string fileName = "Invoice Template.doc";
// Open the template document.
Document doc = new Document(dataDir + fileName);

// Trim trailing and leading whitespaces mail merge values
doc.MailMerge.TrimWhitespaces = false;

// Execute the nested mail merge with regions
doc.MailMerge.ExecuteWithRegions(pizzaDs);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the output to file
doc.Save(dataDir);
