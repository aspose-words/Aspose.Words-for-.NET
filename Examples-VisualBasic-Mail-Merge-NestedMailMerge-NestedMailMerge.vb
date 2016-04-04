' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()

' Create the Dataset and read the XML.
Dim pizzaDs As New DataSet()

' Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
' To see examples of how to set up relations manually check the corresponding documentation of this sample
pizzaDs.ReadXml(dataDir & "CustomerData.xml")

Dim fileName As String = "Invoice Template.doc"
' Open the template document.
Dim doc As New Document(dataDir & fileName)

' Trim trailing and leading whitespaces mail merge values
doc.MailMerge.TrimWhitespaces = False

' Execute the nested mail merge with regions
doc.MailMerge.ExecuteWithRegions(pizzaDs)

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Save the output to file
doc.Save(dataDir)
