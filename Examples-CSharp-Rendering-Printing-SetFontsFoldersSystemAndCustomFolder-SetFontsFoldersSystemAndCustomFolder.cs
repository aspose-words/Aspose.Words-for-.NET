// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

Document doc = new Document(dataDir + "Rendering.doc");

// Retrieve the array of environment-dependent font sources that are searched by default. For example this will contain a "Windows\Fonts\" source on a Windows machines.
// We add this array to a new ArrayList to make adding or removing font entries much easier.
ArrayList fontSources = new ArrayList(FontSettings.GetFontsSources());

// Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);

// Add the custom folder which contains our fonts to the list of existing font sources.
fontSources.Add(folderFontSource);

// Convert the Arraylist of source back into a primitive array of FontSource objects.
FontSourceBase[] updatedFontSources = (FontSourceBase[])fontSources.ToArray(typeof(FontSourceBase));

// Apply the new set of font sources to use.
FontSettings.SetFontsSources(updatedFontSources);
dataDir = dataDir + "Rendering.SetFontsFolders_out_.pdf";
doc.Save(dataDir);
