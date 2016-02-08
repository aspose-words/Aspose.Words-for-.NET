' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_LINQ()

Dim fileName As String = "HelloWorld.doc"
' Load the template document.
Dim doc As New Document(dataDir & fileName)

' Create an instance of sender class to set it's properties.
Dim sender As New Sender() With { _
    .Name = "LINQ Reporting Engine", _
    .Message = "Hello World" _
}

' Create a Reporting Engine.
Dim engine As New ReportingEngine()

' Execute the build report.
engine.BuildReport(doc, sender, "sender")

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

' Save the finished document to disk.
doc.Save(dataDir)
