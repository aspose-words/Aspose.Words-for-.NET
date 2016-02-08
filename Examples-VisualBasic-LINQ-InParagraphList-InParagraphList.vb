' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_LINQ()

Dim fileName As String = "InParagraphList.doc"
' Load the template document.
Dim doc As New Document(dataDir & fileName)

' Create a Reporting Engine.
Dim engine As New ReportingEngine()

' Execute the build report.
engine.BuildReport(doc, Common.GetClients(), "clients")

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

' Save the finished document to disk.
doc.Save(dataDir)
