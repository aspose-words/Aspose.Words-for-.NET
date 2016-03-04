// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
HttpResponse Response = null;
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); ;
// Open an existing document.
Document doc = new Document(dataDir + "MailMerge.ExecuteArray.doc");

// Trim trailing and leading whitespaces mail merge values
doc.MailMerge.TrimWhitespaces = false;

// Fill the fields in the document with user data.
doc.MailMerge.Execute(
    new string[] { "FullName", "Company", "Address", "Address2", "City" },
    new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

dataDir = dataDir + "MailMerge.ExecuteArray_out_.doc";
// Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
doc.Save(Response, dataDir, ContentDisposition.Inline, null);
