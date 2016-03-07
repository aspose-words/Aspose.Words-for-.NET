Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Data
Imports System.Diagnostics

Imports Aspose.Words
Imports Aspose.Words.Reporting
Imports Aspose.Words.MailMerging

Public Class RemoveEmptyRegions
    Public Shared Sub Run()
        ' ExStart:RemoveUnmergedRegions
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
        Dim fileName As String = "TestFile.doc"
        ' Open the document.
        Dim doc As New Document(dataDir & fileName)

        ' Create a dummy data source containing no data.
        Dim data As New DataSet()

        ' ExStart:MailMergeCleanupOptions
        ' Set the appropriate mail merge clean up options to remove any unused regions from the document.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions
        ' doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields
        ' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveStaticFields
        ' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveEmptyParagraphs
        ' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveUnusedFields
        ' ExEnd:MailMergeCleanupOptions

        ' Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
        ' automatically as they are unused.
        doc.MailMerge.ExecuteWithRegions(data)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the output document to disk.
        doc.Save(dataDir)
        ' ExEnd:RemoveUnmergedRegions
        Console.WriteLine(vbNewLine + "Mail merge performed with empty regions successfully." + vbNewLine + "File saved at " + dataDir)
    End Sub
End Class
