' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
Dim fileName As String = "TestFile.Destination.doc"

Dim dstDoc As New Document(dataDir & fileName)
Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

' Set the appended document to start on a new page.
srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Append the source document using the original styles found in the source document.
dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
dstDoc.Save(dataDir)
