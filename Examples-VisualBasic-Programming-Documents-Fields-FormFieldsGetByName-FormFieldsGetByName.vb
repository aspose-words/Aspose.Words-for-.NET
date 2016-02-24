' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

Dim doc As New Document(dataDir & Convert.ToString("FormFields.doc"))
Dim documentFormFields As FormFieldCollection = doc.Range.FormFields

Dim formField1 As FormField = documentFormFields(3)
Dim formField2 As FormField = documentFormFields("Text2")
