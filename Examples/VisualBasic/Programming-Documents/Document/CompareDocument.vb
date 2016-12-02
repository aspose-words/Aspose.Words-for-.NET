Imports System.IO
Imports Aspose.Words
Public Class CompareDocument
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        NormalComparison(dataDir)
        CompareForEqual(dataDir)
    End Sub
    Private Shared Sub NormalComparison(dataDir As String)
        ' ExStart:NormalComparison
        Dim docA As New Document(dataDir & Convert.ToString("TestFile.doc"))
        Dim docB As New Document(dataDir & Convert.ToString("TestFile - Copy.doc"))
        ' DocA now contains changes as revisions. 
        docA.Compare(docB, "user", DateTime.Now)
        ' ExEnd:NormalComparison                     
    End Sub
    Private Shared Sub CompareForEqual(dataDir As String)
        ' ExStart:CompareForEqual
        Dim docA As New Document(dataDir & Convert.ToString("TestFile.doc"))
        Dim docB As New Document(dataDir & Convert.ToString("TestFile - Copy.doc"))
        ' DocA now contains changes as revisions. 
        docA.Compare(docB, "user", DateTime.Now)
        If docA.Revisions.Count = 0 Then
            Console.WriteLine("Documents are equal")
        Else
            Console.WriteLine("Documents are not equal")
        End If
        ' ExEnd:CompareForEqual                     
    End Sub
End Class
