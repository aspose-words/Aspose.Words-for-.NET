Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Markup
Class CheckBoxTypeContentControl
    Public Shared Sub Run()
        ' ExStart:CheckBoxTypeContentControl
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Open the empty document
        Dim doc As New Document()

        Dim builder As New DocumentBuilder(doc)
        Dim SdtCheckBox As New StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)

        ' Insert content control into the document
        builder.InsertNode(SdtCheckBox)
        dataDir = dataDir & Convert.ToString("CheckBoxTypeContentControl_out.docx")

        doc.Save(dataDir, SaveFormat.Docx)
        ' ExEnd:CheckBoxTypeContentControl
        Console.WriteLine(Convert.ToString(vbLf & "CheckBox type content control created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class

