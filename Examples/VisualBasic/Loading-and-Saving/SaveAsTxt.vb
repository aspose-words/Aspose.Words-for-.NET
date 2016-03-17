Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports System.Text

Class SaveAsTxt
    Public Shared Sub Run()
        ' ExStart:SaveAsTxt
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
        dataDir = dataDir & Convert.ToString("Document.ConvertToTxt_out_.txt")
        doc.Save(dataDir)
        ' ExEnd:SaveAsTxt
        Console.WriteLine(Convert.ToString(vbLf & "Document saved as TXT." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class

