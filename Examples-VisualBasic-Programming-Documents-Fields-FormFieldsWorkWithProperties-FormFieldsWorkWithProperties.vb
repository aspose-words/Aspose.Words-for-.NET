' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

Dim doc As New Document(dataDir & Convert.ToString("FormFields.doc"))
Dim formField As FormField = doc.Range.FormFields(3)

If formField.Type.Equals(FieldType.FieldFormTextInput) Then
    formField.Result = "My name is " + formField.Name
End If
