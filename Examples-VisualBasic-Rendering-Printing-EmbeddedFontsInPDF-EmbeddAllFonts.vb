' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))
' Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
' each time a document is rendered.
Dim options As New PdfSaveOptions()
options.EmbedFullFonts = True

Dim outPath As String = dataDir & Convert.ToString("Rendering.EmbedFullFonts_out_.pdf")
' The output PDF will be embedded with all fonts found in the document.
doc.Save(outPath, options)
