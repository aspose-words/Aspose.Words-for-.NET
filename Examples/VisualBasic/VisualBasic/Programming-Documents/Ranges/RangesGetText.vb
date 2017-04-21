Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Public Class RangesGetText
    Public Shared Sub Run()
        ' ExStart:RangesGetText
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithRanges()

        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
        Dim text As String = doc.Range.Text
        ' ExEnd:RangesGetText
        Console.WriteLine(Convert.ToString(vbLf & "Document have following text range ") & text)
    End Sub
End Class
