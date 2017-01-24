Imports System.IO
Imports Aspose.Words
Imports System.Drawing
Imports Aspose.Words.Tables
Class DocumentBuilderInsertBookmark
    Public Shared Sub Run()
        ' ExStart:DocumentBuilderInsertBookmark
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Initialize document.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.StartBookmark("FineBookmark")
        builder.Writeln("This is just a fine bookmark.")
        builder.EndBookmark("FineBookmark")

        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertBookmark_out.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertBookmark
        Console.WriteLine(Convert.ToString(vbLf & "Bookmark using DocumentBuilder inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
