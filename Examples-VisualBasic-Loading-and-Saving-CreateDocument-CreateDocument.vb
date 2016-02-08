' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

' Initialize a Document.
Dim doc As New Document()

' Use a document builder to add content to the document.
Dim builder As New DocumentBuilder(doc)
builder.Writeln("Hello World!")

dataDir = dataDir & Convert.ToString("CreateDocument_out_.docx")
' Save the document to disk.
doc.Save(dataDir)

