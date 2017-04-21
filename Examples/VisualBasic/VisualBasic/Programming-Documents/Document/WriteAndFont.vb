Imports System.IO
Imports Aspose.Words
Imports System.Drawing

Class WriteAndFont
    Public Shared Sub Run()
        ' ExStart:WriteAndFont
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Initialize document.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Specify font formatting before adding text.
        Dim font As Aspose.Words.Font = builder.Font
        font.Size = 16
        font.Bold = True
        font.Color = Color.Blue
        font.Name = "Arial"
        font.Underline = Underline.Dash

        builder.Write("Sample text.")
        dataDir = dataDir & Convert.ToString("WriteAndFont_out.doc")
        doc.Save(dataDir)
        ' ExEnd:WriteAndFont
        Console.WriteLine(Convert.ToString(vbLf & "Formatted text using DocumentBuilder inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
