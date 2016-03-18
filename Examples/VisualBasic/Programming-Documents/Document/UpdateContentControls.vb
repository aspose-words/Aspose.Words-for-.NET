Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Markup
Imports System.Drawing
Imports Aspose.Words.Drawing

Class UpdateContentControls
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        SetCurrentStateOfCheckBox(dataDir)
        'Shows how to modify content controls of type plain text box, drop down list and picture.
        ModifyContentControls(dataDir)
    End Sub
    Public Shared Sub SetCurrentStateOfCheckBox(dataDir As String)
        ' ExStart:SetCurrentStateOfCheckBox
        'Open an existing document
        Dim doc As New Document(dataDir & Convert.ToString("CheckBoxTypeContentControl.docx"))

        Dim builder As New DocumentBuilder(doc)
        'Get the first content control from the document
        Dim SdtCheckBox As StructuredDocumentTag = DirectCast(doc.GetChild(NodeType.StructuredDocumentTag, 0, True), StructuredDocumentTag)

        'StructuredDocumentTag.Checked property gets/sets current state of the Checkbox SDT
        If SdtCheckBox.SdtType = SdtType.Checkbox Then
            SdtCheckBox.Checked = True
        End If

        dataDir = dataDir & Convert.ToString("SetCurrentStateOfCheckBox_out_.docx")
        doc.Save(dataDir)
        ' ExEnd:SetCurrentStateOfCheckBox
        Console.WriteLine(Convert.ToString(vbLf & "Current state fo checkbox setup successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Public Shared Sub ModifyContentControls(dataDir As String)
        ' ExStart:ModifyContentControls
        'Open an existing document
        Dim doc As New Document(dataDir & Convert.ToString("CheckBoxTypeContentControl.docx"))

        For Each sdt As StructuredDocumentTag In doc.GetChildNodes(NodeType.StructuredDocumentTag, True)
            If sdt.SdtType = SdtType.PlainText Then
                sdt.RemoveAllChildren()
                Dim para As Paragraph = TryCast(sdt.AppendChild(New Paragraph(doc)), Paragraph)
                Dim run As New Run(doc, "new text goes here")
                para.AppendChild(run)
            ElseIf sdt.SdtType = SdtType.DropDownList Then
                Dim secondItem As SdtListItem = sdt.ListItems(2)
                sdt.ListItems.SelectedValue = secondItem
            ElseIf sdt.SdtType = SdtType.Picture Then
                Dim shape As Shape = DirectCast(sdt.GetChild(NodeType.Shape, 0, True), Shape)
                If shape.HasImage Then
                    shape.ImageData.SetImage(dataDir & Convert.ToString("Watermark.png"))
                End If
            End If
        Next


        dataDir = dataDir & Convert.ToString("ModifyContentControls_out_.docx")
        doc.Save(dataDir)
        ' ExEnd:ModifyContentControls
        Console.WriteLine(Convert.ToString(vbLf & "Plain text box, drop down list and picture content modified successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
