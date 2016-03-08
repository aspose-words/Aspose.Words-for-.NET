' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

' Load the documents which store the shapes we want to render.
Dim doc As New Document(dataDir & Convert.ToString("TestFile RenderShape.doc"))
Hyphenation.RegisterDictionary("en-US", dataDir & Convert.ToString("hyph_en_US.dic"))
Hyphenation.RegisterDictionary("de-CH", dataDir & Convert.ToString("hyph_de_CH.dic"))

dataDir = dataDir & Convert.ToString("HyphenateWordsOfLanguages_out_.pdf")
doc.Save(dataDir)
