' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))

' Retrieve the array of environment-dependent font sources that are searched by default. For example this will contain a "Windows\Fonts\" source on a Windows machines.
' We add this array to a new ArrayList to make adding or removing font entries much easier.
Dim fontSources As New ArrayList(FontSettings.GetFontsSources())

' Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
Dim folderFontSource As New FolderFontSource("C:\MyFonts\", True)

' Add the custom folder which contains our fonts to the list of existing font sources.
fontSources.Add(folderFontSource)

' Convert the Arraylist of source back into a primitive array of FontSource objects.
Dim updatedFontSources As FontSourceBase() = DirectCast(fontSources.ToArray(GetType(FontSourceBase)), FontSourceBase())

' Apply the new set of font sources to use.
FontSettings.SetFontsSources(updatedFontSources)
dataDir = dataDir & Convert.ToString("Rendering.SetFontsFolders_out_.pdf")
doc.Save(dataDir)
