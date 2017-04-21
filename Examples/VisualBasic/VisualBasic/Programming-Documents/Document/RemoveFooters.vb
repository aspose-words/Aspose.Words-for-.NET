Imports Aspose.Words
Public Class RemoveFooters
    Public Shared Sub Run()
        ' ExStart:RemoveFooters
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()

        Dim doc As New Document(dataDir & Convert.ToString("HeaderFooter.RemoveFooters.doc"))

        For Each section As Section In doc
            ' Up to three different footers are possible in a section (for first, even and odd pages).
            ' We check and delete all of them.
            Dim footer As HeaderFooter

            footer = section.HeadersFooters(HeaderFooterType.FooterFirst)
            If footer IsNot Nothing Then
                footer.Remove()
            End If

            ' Primary footer is the footer used for odd pages.
            footer = section.HeadersFooters(HeaderFooterType.FooterPrimary)
            If footer IsNot Nothing Then
                footer.Remove()
            End If

            footer = section.HeadersFooters(HeaderFooterType.FooterEven)
            If footer IsNot Nothing Then
                footer.Remove()
            End If
        Next
        dataDir = dataDir & Convert.ToString("HeaderFooter.RemoveFooters_out.doc")

        ' Save the document.
        doc.Save(dataDir)
        ' ExEnd:RemoveFooters
        Console.WriteLine(Convert.ToString(vbLf & "All footers from all sections deleted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
