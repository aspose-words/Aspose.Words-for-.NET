' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

' Loads encrypted document.
Dim doc As New Document(dataDir & Convert.ToString("LoadEncrypted.docx"), New LoadOptions("aspose"))

