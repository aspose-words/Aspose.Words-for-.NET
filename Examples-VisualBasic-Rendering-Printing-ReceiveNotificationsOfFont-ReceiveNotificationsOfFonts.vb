' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()



Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))
' We can choose the default font to use in the case of any missing fonts.
FontSettings.DefaultFontName = "Arial"
' For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
' find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
' font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
FontSettings.SetFontsFolder(String.Empty, False)

' Create a new class implementing IWarningCallback which collect any warnings produced during document save.
Dim callback As New HandleDocumentWarnings()

doc.WarningCallback = callback
Dim path As String = dataDir & Convert.ToString("Rendering.MissingFontNotification_out_.pdf")
' Pass the save options along with the save path to the save method.
doc.Save(path)
