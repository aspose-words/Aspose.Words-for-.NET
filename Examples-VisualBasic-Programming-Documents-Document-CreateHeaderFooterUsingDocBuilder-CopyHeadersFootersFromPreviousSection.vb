' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
''' <summary>
''' Clones and copies headers/footers form the previous section to the specified section.
''' </summary>
Private Shared Sub CopyHeadersFootersFromPreviousSection(section As Section)
    Dim previousSection As Section = DirectCast(section.PreviousSibling, Section)

    If previousSection Is Nothing Then
        Return
    End If

    section.HeadersFooters.Clear()

    For Each headerFooter As HeaderFooter In previousSection.HeadersFooters
        section.HeadersFooters.Add(headerFooter.Clone(True))
    Next
End Sub
