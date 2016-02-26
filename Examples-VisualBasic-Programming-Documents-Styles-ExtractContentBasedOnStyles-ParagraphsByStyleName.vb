' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Public Shared Function ParagraphsByStyleName(ByVal doc As Document, ByVal styleName As String) As ArrayList
    ' Create an array to collect paragraphs of the specified style.
    Dim paragraphsWithStyle As New ArrayList()
    ' Get all paragraphs from the document.
    Dim paragraphs As NodeCollection = doc.GetChildNodes(NodeType.Paragraph, True)
    ' Look through all paragraphs to find those with the specified style.
    For Each paragraph As Paragraph In paragraphs
        If paragraph.ParagraphFormat.Style.Name = styleName Then
            paragraphsWithStyle.Add(paragraph)
        End If
    Next paragraph
    Return paragraphsWithStyle
End Function
