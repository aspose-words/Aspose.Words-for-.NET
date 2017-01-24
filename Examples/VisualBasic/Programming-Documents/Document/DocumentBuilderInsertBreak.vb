Imports System.IO
Imports Aspose.Words
Imports System.Drawing
Imports Aspose.Words.Tables

Class DocumentBuilderInsertBreak
    Public Shared Sub Run()
        ' ExStart:DocumentBuilderInsertBreak
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Initialize document.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.Writeln("This is page 1.")
        builder.InsertBreak(BreakType.PageBreak)

        builder.Writeln("This is page 2.")
        builder.InsertBreak(BreakType.PageBreak)

        builder.Writeln("This is page 3.")
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertBreak_out.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertBreak
        Console.WriteLine(Convert.ToString(vbLf & "Page breaks inserted into a document using DocumentBuilder." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
