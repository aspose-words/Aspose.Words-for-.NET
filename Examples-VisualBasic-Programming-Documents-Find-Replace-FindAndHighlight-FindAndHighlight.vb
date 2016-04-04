' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_FindAndReplace()
Dim fileName As String = "TestFile.doc"

Dim doc As New Document(dataDir & fileName)

' We want the "your document" phrase to be highlighted.
Dim regex As New Regex("your document", RegexOptions.IgnoreCase)
doc.Range.Replace(regex, New ReplaceEvaluatorFindAndHighlight(), False)

dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
' Save the output document.
doc.Save(dataDir)
