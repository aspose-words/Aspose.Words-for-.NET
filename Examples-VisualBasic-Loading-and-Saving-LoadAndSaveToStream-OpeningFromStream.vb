' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

' Open the stream. Read only access is enough for Aspose.Words to load a document.
Dim stream As Stream = File.OpenRead(dataDir & "Document.doc")

' Load the entire document into memory.
Dim doc As New Document(stream)

' You can close the stream now, it is no longer needed because the document is in memory.
stream.Close()
