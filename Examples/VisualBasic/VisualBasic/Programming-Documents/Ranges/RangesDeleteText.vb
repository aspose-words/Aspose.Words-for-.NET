Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Public Class RangesDeleteText
    Public Shared Sub Run()
        ' ExStart:RangesDeleteText
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithRanges()

        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
        doc.Sections(0).Range.Delete()
        ' ExEnd:RangesDeleteText
        Console.WriteLine(vbLf & "All characters of a range deleted successfully.")
    End Sub
End Class
