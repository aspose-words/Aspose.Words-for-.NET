' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim Response As HttpResponse = Nothing
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()


' Open an existing document.
Dim doc As New Document(dataDir & Convert.ToString("MailMerge.ExecuteArray.doc"))

' Trim trailing and leading whitespaces mail merge values
doc.MailMerge.TrimWhitespaces = False

' Fill the fields in the document with user data.
doc.MailMerge.Execute(New String() {"FullName", "Company", "Address", "Address2", "City"}, New Object() {"James Bond", "MI5 Headquarters", "Milbank", "", "London"})

dataDir = dataDir & Convert.ToString("MailMerge.ExecuteArray_out_.doc")
' Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
doc.Save(Response, dataDir, ContentDisposition.Inline, Nothing)
