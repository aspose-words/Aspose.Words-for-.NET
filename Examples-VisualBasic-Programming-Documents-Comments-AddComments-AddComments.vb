' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithComments()
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)
builder.Write("Some text is added.")

Dim comment As New Comment(doc, "Awais Hafeez", "AH", DateTime.Today)
builder.CurrentParagraph.AppendChild(comment)
comment.Paragraphs.Add(New Paragraph(doc))
comment.FirstParagraph.Runs.Add(New Run(doc, "Comment text."))

dataDir = dataDir & Convert.ToString("Comments_out_.doc")
' Save the document.
doc.Save(dataDir)
