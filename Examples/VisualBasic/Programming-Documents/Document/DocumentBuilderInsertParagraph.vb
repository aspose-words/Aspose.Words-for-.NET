Imports System.IO
Imports Aspose.Words
Imports System.Drawing

Class DocumentBuilderInsertParagraph
    Public Shared Sub Run()
        ' ExStart:DocumentBuilderInsertParagraph
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Initialize document.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Specify font formatting
        Dim font As Aspose.Words.Font = builder.Font
        font.Size = 16
        font.Bold = True
        font.Color = System.Drawing.Color.Blue
        font.Name = "Arial"
        font.Underline = Underline.Dash

        ' Specify paragraph formatting
        Dim paragraphFormat As ParagraphFormat = builder.ParagraphFormat
        paragraphFormat.FirstLineIndent = 8
        paragraphFormat.Alignment = ParagraphAlignment.Justify
        paragraphFormat.KeepTogether = True

        builder.Writeln("A whole paragraph.")
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertParagraph_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertParagraph
        Console.WriteLine(Convert.ToString(vbLf & "Paragraph inserted successfully into the document using DocumentBuilder." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class

