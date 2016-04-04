' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

' Load the document from disk.
Dim doc As New Document(dataDir & "Template.doc")

dataDir = dataDir & "Template_out_.pdf"
' Save the document in PDF format.
doc.Save(dataDir)
