' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()

' Load the document from disk.
Dim doc As New Document(dataDir & Convert.ToString("TestFile.doc"))

Dim clone As Document = doc.Clone()

dataDir = dataDir & Convert.ToString("TestFile_clone_out_.doc")

' Save the document to disk.
clone.Save(dataDir)
