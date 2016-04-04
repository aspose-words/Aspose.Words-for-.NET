// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
string fileName = "Template.doc";
// Load the template document.
Document doc = new Document(dataDir + fileName);

// Setup mail merge event handler to do the custom work.
doc.MailMerge.FieldMergingCallback = new HandleMergeField();

// Trim trailing and leading whitespaces mail merge values
doc.MailMerge.TrimWhitespaces = false;

// This is the data for mail merge.
String[] fieldNames = new String[] {"RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
    "Subject", "Body", "Urgent", "ForReview", "PleaseComment"};
Object[] fieldValues = new Object[] {"Josh", "Jenny", "123456789", "", "Hello",
    "<b>HTML Body Test message 1</b>", true, false, true};

// Execute the mail merge.
doc.MailMerge.Execute(fieldNames, fieldValues);

dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
// Save the finished document.
doc.Save(dataDir);
