' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_QuickStart()
Dim fileName As String = "Document.doc"
' Load the document from the absolute path on disk.
Dim doc As New Document(dataDir & fileName)

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Save the document as DOCX document.");
doc.Save(dataDir)
