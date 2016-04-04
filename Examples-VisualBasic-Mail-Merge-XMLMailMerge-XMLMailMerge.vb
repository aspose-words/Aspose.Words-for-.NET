' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()

' Create the Dataset and read the XML.
Dim customersDs As New DataSet()
customersDs.ReadXml(dataDir & "Customers.xml")

Dim fileName As String = "TestFile.doc"
' Open a template document.
Dim doc As New Document(dataDir & fileName)

' Execute mail merge to fill the template with data from XML using DataTable.
doc.MailMerge.Execute(customersDs.Tables("Customer"))

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Save the output document.
doc.Save(dataDir)
