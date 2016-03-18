Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Markup
Imports System.Drawing
Class RichTextBoxContentControl
    Public Shared Sub Run()
        ' ExStart:RichTextBoxContentControl
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        Dim doc As New Document()
        Dim sdtRichText As New StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)

        Dim para As New Paragraph(doc)
        Dim run As New Run(doc)
        run.Text = "Hello World"
        run.Font.Color = Color.Green
        para.Runs.Add(run)
        sdtRichText.ChildNodes.Add(para)
        doc.FirstSection.Body.AppendChild(sdtRichText)

        dataDir = dataDir & Convert.ToString("RichTextBoxContentControl_out_.docx")
        doc.Save(dataDir)
        ' ExEnd:RichTextBoxContentControl
        Console.WriteLine(Convert.ToString(vbLf & "Rich text box type content control created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
