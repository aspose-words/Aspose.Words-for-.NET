Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Public Class ExtractContentBasedOnStyles
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithStyles()

        ' Open the document.
        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Define style names as they are specified in the Word document.
        Const paraStyle As String = "Heading 1"
        Const runStyle As String = "Intense Emphasis"

        ' Collect paragraphs with defined styles. 
        ' Show the number of collected paragraphs and display the text of this paragraphs.
        Dim paragraphs As ArrayList = ParagraphsByStyleName(doc, paraStyle)
        Console.WriteLine(String.Format("Paragraphs with ""{0}"" styles ({1}):", paraStyle, paragraphs.Count))
        For Each paragraph As Paragraph In paragraphs
            Console.Write(paragraph.ToString(SaveFormat.Text))
        Next paragraph

        ' Collect runs with defined styles. 
        ' Show the number of collected runs and display the text of this runs.
        Dim runs As ArrayList = RunsByStyleName(doc, runStyle)
        Console.WriteLine(String.Format(Constants.vbLf & "Runs with ""{0}"" styles ({1}):", runStyle, runs.Count))
        For Each run As Run In runs
            Console.WriteLine(run.Range.Text)
        Next run

        Console.WriteLine(vbNewLine & "Extracted contents based on styles successfully.")
    End Sub

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

    Public Shared Function RunsByStyleName(ByVal doc As Document, ByVal styleName As String) As ArrayList
        ' Create an array to collect runs of the specified style.
        Dim runsWithStyle As New ArrayList()
        ' Get all runs from the document.
        Dim runs As NodeCollection = doc.GetChildNodes(NodeType.Run, True)
        ' Look through all runs to find those with the specified style.
        For Each run As Run In runs
            If run.Font.Style.Name = styleName Then
                runsWithStyle.Add(run)
            End If
        Next run
        Return runsWithStyle
    End Function
End Class
