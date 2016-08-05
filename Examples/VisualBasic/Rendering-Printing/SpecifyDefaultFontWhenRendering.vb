Imports System.IO
Imports System.Reflection
Imports Aspose.Words.Fonts
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class SpecifyDefaultFontWhenRendering
    Public Shared Sub Run()
        ' ExStart:SpecifyDefaultFontWhenRendering
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
        Dim FontSettings As New FontSettings()
        Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))

        ' If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
        FontSettings.DefaultFontName = "Arial Unicode MS"
        ' Set font settings
        doc.FontSettings = FontSettings
        dataDir = dataDir & Convert.ToString("Rendering.SetDefaultFont_out_.pdf")
        ' Now the set default font is used in place of any missing fonts during any rendering calls.
        doc.Save(dataDir)
        ' ExEnd:SpecifyDefaultFontWhenRendering 
        Console.WriteLine(Convert.ToString(vbLf & "Default font is setup during rendering." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
