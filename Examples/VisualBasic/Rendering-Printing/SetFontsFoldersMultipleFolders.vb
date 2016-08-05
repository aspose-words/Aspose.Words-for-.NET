Imports System.IO
Imports System.Reflection
Imports Aspose.Words.Fonts
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class SetFontsFoldersMultipleFolders
    Public Shared Sub Run()
        ' ExStart:SetFontsFoldersMultipleFolders
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
        Dim FontSettings As New FontSettings()
        Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))
        ' Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        ' fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        ' FontSettings.SetFontSources instead.
        FontSettings.SetFontsFolders(New String() {"C:\MyFonts\", "D:\Misc\Fonts\"}, True)
        ' Set font settings
        doc.FontSettings = FontSettings
        dataDir = dataDir & Convert.ToString("Rendering.SetFontsFolders_out_.pdf")
        doc.Save(dataDir)
        ' ExEnd:SetFontsFoldersMultipleFolders           
    End Sub
End Class
