' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
Dim fileName As String = "TestFile.doc"
' Open the document.
Dim doc As New Document(dataDir & fileName)

' Create a dummy data source containing no data.
Dim data As New DataSet()

' Set the appropriate mail merge clean up options to remove any unused regions from the document.
doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions
' doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields
' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveStaticFields
' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveEmptyParagraphs
' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveUnusedFields

' Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
' automatically as they are unused.
doc.MailMerge.ExecuteWithRegions(data)

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Save the output document to disk.
doc.Save(dataDir)
