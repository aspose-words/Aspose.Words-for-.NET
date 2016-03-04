// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

Document doc = new Document(dataDir + "Rendering.doc");
// We can choose the default font to use in the case of any missing fonts.
FontSettings.DefaultFontName = "Arial";
// For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
// find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
// font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
FontSettings.SetFontsFolder(string.Empty, false);

// Create a new class implementing IWarningCallback which collect any warnings produced during document save.
HandleDocumentWarnings callback = new HandleDocumentWarnings();

doc.WarningCallback = callback;
string path = dataDir + "Rendering.MissingFontNotification_out_.pdf";
// Pass the save options along with the save path to the save method.
doc.Save(path);
