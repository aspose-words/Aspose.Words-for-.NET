Imports System.Collections
Imports System.IO
Imports System.Text.RegularExpressions
Imports Aspose.Words
Imports Aspose.Words.Fields
Imports Aspose.Words.Replacing
Class ReplaceWithRegex
    Public Shared Sub Run()
        ' ExStart:ReplaceWithRegex
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_FindAndReplace()

        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
        doc.Range.Replace(New Regex("[s|m]ad"), "bad", New FindReplaceOptions(FindReplaceDirection.Forward))

        dataDir = dataDir & Convert.ToString("ReplaceWithRegex_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:ReplaceWithRegex
        Console.WriteLine(Convert.ToString(vbLf & "Text replaced with regex successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
