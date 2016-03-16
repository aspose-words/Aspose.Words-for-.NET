Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Imports Aspose.Words.Fields
Imports Aspose.Words.Tables
Class DocumentBuilderInsertImage
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        InsertInlineImage(dataDir)
        InsertFloatingImage(dataDir)
    End Sub
    Public Shared Sub InsertInlineImage(dataDir As String)
        ' ExStart:DocumentBuilderInsertInlineImage
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.InsertImage(dataDir & Convert.ToString("Watermark.png"))
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertInlineImage_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertInlineImage
        Console.WriteLine(Convert.ToString(vbLf & "Inline image using DocumentBuilder inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub InsertFloatingImage(dataDir As String)
        ' ExStart:DocumentBuilderInsertFloatingImage
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.InsertImage(dataDir & Convert.ToString("Watermark.png"), RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100, 200, _
            100, WrapType.Square)
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertFloatingImage_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertFloatingImage
        Console.WriteLine(Convert.ToString(vbLf & "Inline image using DocumentBuilder inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class

