Imports System.IO
Imports System.Reflection
Imports Aspose.Words.Fonts
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class SetTrueTypeFontsFolder
    Public Shared Sub Run()
        ' ExStart:SetTrueTypeFontsFolder
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
        Dim FontSettings As New FontSettings()
        Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))

        ' Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        ' fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        ' FontSettings.SetFontSources instead.
        FontSettings.SetFontsFolder("C:\MyFonts\", False)
        ' Set font settings
        doc.FontSettings = FontSettings
        dataDir = dataDir & Convert.ToString("Rendering.SetFontsFolder_out_.pdf")
        doc.Save(dataDir)
        ' ExEnd:SetTrueTypeFontsFolder
        Console.WriteLine(Convert.ToString(vbLf & "True type fonts folder setup successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
