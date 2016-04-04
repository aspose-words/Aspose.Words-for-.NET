' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
Dim fileName As String = "Template.doc"
' Load the template document.
Dim doc As New Document(dataDir & fileName)

' Setup mail merge event handler to do the custom work.
doc.MailMerge.FieldMergingCallback = New HandleMergeField()

' Trim trailing and leading whitespaces mail merge values
doc.MailMerge.TrimWhitespaces = False

' This is the data for mail merge.
Dim fieldNames() As String = {"RecipientName", "SenderName", "FaxNumber", "PhoneNumber", "Subject", "Body", "Urgent", "ForReview", "PleaseComment"}
Dim fieldValues() As Object = {"Josh", "Jenny", "123456789", "", "Hello", "<b>HTML Body Test message 1</b>", True, False, True}

' Execute the mail merge.
doc.MailMerge.Execute(fieldNames, fieldValues)

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Save the finished document.
doc.Save(dataDir)
