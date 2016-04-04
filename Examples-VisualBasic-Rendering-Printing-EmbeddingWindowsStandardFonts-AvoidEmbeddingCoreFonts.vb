' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))
' To disable embedding of core fonts and subsuite PDF type 1 fonts set UseCoreFonts to true.
Dim options As New PdfSaveOptions()
options.UseCoreFonts = True

Dim outPath As String = dataDir & Convert.ToString("Rendering.DisableEmbedWindowsFonts_out_.pdf")
' The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
doc.Save(outPath)
