// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
Document doc = new Document(dataDir + "MailMerge.MergeImage.doc");

// Set up the event handler for image fields.
doc.MailMerge.FieldMergingCallback = new HandleMergeImageFieldFromBlob();

// Open a database connection.
string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + RunExamples.GetDataDir_Database()+"Northwind.mdb";
OleDbConnection conn = new OleDbConnection(connString);
conn.Open();

// Open the data reader. It needs to be in the normal mode that reads all record at once.
OleDbCommand cmd = new OleDbCommand("SELECT * FROM Employees", conn);
IDataReader dataReader = cmd.ExecuteReader();

// Perform mail merge.
doc.MailMerge.ExecuteWithRegions(dataReader, "Employees");

// Close the database.
conn.Close();
dataDir = dataDir + "MailMerge.MergeImage_out_.doc";
doc.Save(dataDir);
