' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

' Specify your document name here.
Dim doc As New Document(dataDir & Convert.ToString("RenameMergeFields.doc"))

' Select all field start nodes so we can find the merge fields.
Dim fieldStarts As NodeCollection = doc.GetChildNodes(NodeType.FieldStart, True)
For Each fieldStart As FieldStart In fieldStarts
    If fieldStart.FieldType.Equals(FieldType.FieldMergeField) Then
        Dim mergeField As New MergeField(fieldStart)
        mergeField.Name = mergeField.Name & Convert.ToString("_Renamed")
    End If
Next

dataDir = dataDir & Convert.ToString("RenameMergeFields_out_.doc")
doc.Save(dataDir)
