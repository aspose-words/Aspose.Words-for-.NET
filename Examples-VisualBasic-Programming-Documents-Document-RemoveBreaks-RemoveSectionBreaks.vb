' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Private Shared Sub RemoveSectionBreaks(ByVal doc As Document)
    ' Loop through all sections starting from the section that precedes the last one 
    ' and moving to the first section.
    For i As Integer = doc.Sections.Count - 2 To 0 Step -1
        ' Copy the content of the current section to the beginning of the last section.
        doc.LastSection.PrependContent(doc.Sections(i))
        ' Remove the copied section.
        doc.Sections(i).Remove()
    Next i
End Sub
