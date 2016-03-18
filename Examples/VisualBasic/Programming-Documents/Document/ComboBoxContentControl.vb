Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Markup
Imports System.Drawing

Class ComboBoxContentControl
    Public Shared Sub Run()
        ' ExStart:ComboBoxContentControl
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        Dim doc As New Document()
        Dim sdt As New StructuredDocumentTag(doc, SdtType.ComboBox, MarkupLevel.Block)

        sdt.ListItems.Add(New SdtListItem("Choose an item", "-1"))
        sdt.ListItems.Add(New SdtListItem("Item 1", "1"))
        sdt.ListItems.Add(New SdtListItem("Item 2", "2"))
        doc.FirstSection.Body.AppendChild(sdt)

        dataDir = dataDir & Convert.ToString("ComboBoxContentControl_out_.docx")
        doc.Save(dataDir)
        ' ExEnd:ComboBoxContentControl
        Console.WriteLine(Convert.ToString(vbLf & "Combo box type content control created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
