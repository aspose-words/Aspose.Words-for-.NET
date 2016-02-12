' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithSections()

Dim srcDoc As New Document(dataDir & Convert.ToString("Document.doc"))
Dim dstDoc As New Document()

Dim sourceSection As Section = srcDoc.Sections(0)
Dim newSection As Section = DirectCast(dstDoc.ImportNode(sourceSection, True), Section)
dstDoc.Sections.Add(newSection)
dataDir = dataDir & Convert.ToString("Document.Copy_out_.doc")
dstDoc.Save(dataDir)
