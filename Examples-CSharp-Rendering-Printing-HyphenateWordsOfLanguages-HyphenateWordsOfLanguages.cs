// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

// Load the documents which store the shapes we want to render.
Document doc = new Document(dataDir + "TestFile RenderShape.doc");
Hyphenation.RegisterDictionary("en-US", dataDir + @"hyph_en_US.dic");
Hyphenation.RegisterDictionary("de-CH", dataDir + @"hyph_de_CH.dic");

dataDir = dataDir + "HyphenateWordsOfLanguages_out_.pdf";
doc.Save(dataDir);
