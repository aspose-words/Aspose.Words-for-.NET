// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithFields();

// Create a blank document.
Document doc = new Document();
DocumentBuilder b = new DocumentBuilder(doc);
b.InsertField("MERGEFIELD Date");

// Store the current culture so it can be set back once mail merge is complete.
CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
// Set to German language so dates and numbers are formatted using this culture during mail merge.
Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

// Execute mail merge.
doc.MailMerge.Execute(new string[] { "Date" }, new object[] { DateTime.Now });

// Restore the original culture.
Thread.CurrentThread.CurrentCulture = currentCulture;
doc.Save(dataDir + "Field.ChangeLocale_out_.doc");
