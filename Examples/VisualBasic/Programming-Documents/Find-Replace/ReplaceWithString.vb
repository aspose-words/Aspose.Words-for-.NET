Imports System.Collections
Imports System.IO
Imports System.Text.RegularExpressions
Imports Aspose.Words
Imports Aspose.Words.Fields

Class ReplaceWithString
    Public Shared Sub Run()
        ' ExStart:ReplaceWithString
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_FindAndReplace()
        Dim fileName As String = "Document.doc"

        Dim doc As New Document(dataDir & fileName)
        doc.Range.Replace("sad", "bad", False, True)

        dataDir = dataDir & Convert.ToString("ReplaceWithString_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:ReplaceWithString
        Console.WriteLine(Convert.ToString(vbLf & "Text replaced with string successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class

