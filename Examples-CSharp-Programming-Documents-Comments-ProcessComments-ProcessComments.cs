// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithComments();
string fileName = "TestFile.doc";

// Open the document.
Document doc = new Document(dataDir + fileName);

// Extract the information about the comments of all the authors.
foreach (string comment in ExtractComments(doc))
    Console.Write(comment);

// Remove comments by the "pm" author.
RemoveComments(doc, "pm");
Console.WriteLine("Comments from \"pm\" are removed!");

// Extract the information about the comments of the "ks" author.
foreach (string comment in ExtractComments(doc, "ks"))
    Console.Write(comment);

// Remove all comments.
RemoveComments(doc);
Console.WriteLine("All comments are removed!");

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the document.
doc.Save(dataDir);
