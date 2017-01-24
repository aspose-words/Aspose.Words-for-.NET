Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Layout
Class ChangeStyleOfTOCLevel
    Public Shared Sub Run()
        ' ExStart:ChangeStyleOfTOCLevel
        Dim doc As New Document()
        ' Retrieve the style used for the first level of the TOC and change the formatting of the style.
        doc.Styles(StyleIdentifier.Toc1).Font.Bold = True
        ' ExEnd:ChangeStyleOfTOCLevel 
        Console.WriteLine(vbLf & "TOC level style changed successfully.")
    End Sub
End Class
