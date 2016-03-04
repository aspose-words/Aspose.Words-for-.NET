// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

Document doc = new Document(dataDir + "Rendering.doc");

// Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
// fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
// FontSettings.SetFontSources instead.
FontSettings.SetFontsFolders(new string[] { @"C:\MyFonts\", @"D:\Misc\Fonts\" }, true);
dataDir = dataDir + "Rendering.SetFontsFolders_out_.pdf";
doc.Save(dataDir);
