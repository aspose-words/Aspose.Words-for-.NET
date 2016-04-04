' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

' Load the documents which store the shapes we want to render.
Dim doc As New Document(dataDir & Convert.ToString("TestFile RenderShape.doc"))
Dim stream As Stream = File.OpenRead(dataDir & Convert.ToString("hyph_de_CH.dic"))
Hyphenation.RegisterDictionary("de-CH", stream)

dataDir = dataDir & Convert.ToString("LoadHyphenationDictionaryForLanguage_out_.pdf")
doc.Save(dataDir)
