// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

// Load the documents which store the shapes we want to render.
Document doc = new Document(dataDir + "TestFile RenderShape.doc");
Stream stream = File.OpenRead(dataDir + @"hyph_de_CH.dic");
Hyphenation.RegisterDictionary("de-CH", stream);

dataDir = dataDir + "LoadHyphenationDictionaryForLanguage_out_.pdf";
doc.Save(dataDir);
