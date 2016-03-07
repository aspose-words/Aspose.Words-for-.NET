// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
string fileName = "TestFile.LINQ.doc";
// Open the template document.
Document doc = new Document(dataDir + fileName);

// Fill the document with data from our data sources.
// Using mail merge regions for populating the order items table is required
// because it allows the region to be repeated in the document for each order item.
doc.MailMerge.ExecuteWithRegions(orderItemsDataSource);

// The standard mail merge without regions is used for the delivery address.
doc.MailMerge.Execute(deliveryDataSource);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the output document.
doc.Save(dataDir);
