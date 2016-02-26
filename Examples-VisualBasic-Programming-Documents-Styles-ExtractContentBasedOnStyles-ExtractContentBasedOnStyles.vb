' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithStyles()
Dim fileName As String = "TestFile.doc"
' Open the document.
Dim doc As New Document(dataDir & fileName)

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
