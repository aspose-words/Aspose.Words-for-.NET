Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Public Class RemoveBreaks
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        Dim fileName As String = "TestFile.doc"
        ' Open the document.
        Dim doc As New Document(dataDir & fileName)

        ' Remove the page and section breaks from the document.
        ' In Aspose.Words section breaks are represented as separate Section nodes in the document.
        ' To remove these separate sections the sections are combined.
        RemovePageBreaks(doc)
        RemoveSectionBreaks(doc)

        dataDir = dataDir + RunExamples.GetOutputFilePath(fileName)
        ' Save the document.
        doc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Removed breaks from the document successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

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
End Class
