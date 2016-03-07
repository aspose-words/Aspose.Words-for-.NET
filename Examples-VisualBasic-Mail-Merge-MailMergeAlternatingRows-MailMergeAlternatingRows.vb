' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()


Dim doc As New Document(dataDir & Convert.ToString("MailMerge.AlternatingRows.doc"))

' Add a handler for the MergeField event.
doc.MailMerge.FieldMergingCallback = New HandleMergeFieldAlternatingRows()

' Execute mail merge with regions.
Dim dataTable As DataTable = GetSuppliersDataTable()
doc.MailMerge.ExecuteWithRegions(dataTable)
dataDir = dataDir & Convert.ToString("MailMerge.AlternatingRows_out_.doc")
doc.Save(dataDir)
