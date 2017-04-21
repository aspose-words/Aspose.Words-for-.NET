Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Layout

Class ChangeTOCTabStops
    Public Shared Sub Run()
        ' ExStart:ChangeTOCTabStops
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithStyles()

        Dim fileName As String = "Document.TableOfContents.doc"
        ' Open the document.
        Dim doc As New Document(dataDir & fileName)

        ' Iterate through all paragraphs in the document
        For Each para As Paragraph In doc.GetChildNodes(NodeType.Paragraph, True)
            ' Check if this paragraph is formatted using the TOC result based styles. This is any style between TOC and TOC9.
            If para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 AndAlso para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9 Then
                ' Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                Dim tab As TabStop = para.ParagraphFormat.TabStops(0)
                ' Remove the old tab from the collection.
                para.ParagraphFormat.TabStops.RemoveByPosition(tab.Position)
                ' Insert a new tab using the same properties but at a modified position. 
                ' We could also change the separators used (dots) by passing a different Leader type
                para.ParagraphFormat.TabStops.Add(tab.Position - 50, tab.Alignment, tab.Leader)
            End If
        Next

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        doc.Save(dataDir)
        ' ExEnd:ChangeTOCTabStops 
        Console.WriteLine(Convert.ToString(vbLf & "Position of the right tab stop in TOC related paragraphs modified successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class

