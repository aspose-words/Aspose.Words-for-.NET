' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections()

Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))
Dim section As Section = doc.Sections(0)
section.ClearHeadersFooters()
