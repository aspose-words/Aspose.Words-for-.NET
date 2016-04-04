' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
Dim doc As New Document(dataDir & Convert.ToString("MailMerge.MergeImage.doc"))

' Set up the event handler for image fields.
doc.MailMerge.FieldMergingCallback = New HandleMergeImageFieldFromBlob()

' Open a database connection.
Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + RunExamples.GetDataDir_Database() + "Northwind.mdb"
Dim conn As New OleDbConnection(connString)
conn.Open()

' Open the data reader. It needs to be in the normal mode that reads all record at once.
Dim cmd As New OleDbCommand("SELECT * FROM Employees", conn)
Dim dataReader As IDataReader = cmd.ExecuteReader()

' Perform mail merge.
doc.MailMerge.ExecuteWithRegions(dataReader, "Employees")

' Close the database.
conn.Close()
dataDir = dataDir & Convert.ToString("MailMerge.MergeImage_out_.doc")
doc.Save(dataDir)
