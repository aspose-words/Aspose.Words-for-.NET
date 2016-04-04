' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)
builder.InsertField("MERGEFIELD MyFieldName \* MERGEFORMAT")
dataDir = dataDir & Convert.ToString("InsertField_out_.docx")
doc.Save(dataDir)
