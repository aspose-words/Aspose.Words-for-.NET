' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Private Shared Sub RemovePageBreaks(ByVal doc As Document)
    ' Retrieve all paragraphs in the document.
    Dim paragraphs As NodeCollection = doc.GetChildNodes(NodeType.Paragraph, True)

    ' Iterate through all paragraphs
    For Each para As Paragraph In paragraphs
        ' If the paragraph has a page break before set then clear it.
        If para.ParagraphFormat.PageBreakBefore Then
            para.ParagraphFormat.PageBreakBefore = False
        End If

        ' Check all runs in the paragraph for page breaks and remove them.
        For Each run As Run In para.Runs
            If run.Text.Contains(ControlChar.PageBreak) Then
                run.Text = run.Text.Replace(ControlChar.PageBreak, String.Empty)
            End If
        Next run

    Next para

End Sub
