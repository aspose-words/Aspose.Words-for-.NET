' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

' The encoding of the text file is automatically detected.
Dim doc As New Document(dataDir & Convert.ToString("LoadTxt.txt"))

dataDir = dataDir & "LoadTxt_out_.docx"
' Save as any Aspose.Words supported format, such as DOCX.
doc.Save(dataDir)
