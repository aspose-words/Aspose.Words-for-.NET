Imports Microsoft.VisualBasic
Imports Aspose.Words

Public Class CloningDocument
    Public Shared Sub Run()
        ' ExStart:CloningDocument
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()

        ' Load the document from disk.
        Dim doc As New Document(dataDir & Convert.ToString("TestFile.doc"))

        Dim clone As Document = doc.Clone()

        dataDir = dataDir & Convert.ToString("TestFile_clone_out.doc")

        ' Save the document to disk.
        clone.Save(dataDir)
        ' ExEnd:CloningDocument
        Console.WriteLine(Convert.ToString(vbLf & "Document cloned successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
