Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Layout
Imports Aspose.Words.Rendering
Public Class LoadHyphenationDictionaryForLanguage
    Public Shared Sub Run()
        ' ExStart:LoadHyphenationDictionaryForLan
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        ' Load the documents which store the shapes we want to render.
        Dim doc As New Document(dataDir & Convert.ToString("TestFile RenderShape.doc"))
        Dim stream As Stream = File.OpenRead(dataDir & Convert.ToString("hyph_de_CH.dic"))
        Hyphenation.RegisterDictionary("de-CH", stream)

        dataDir = dataDir & Convert.ToString("LoadHyphenationDictionaryForLanguage_out_.pdf")
        doc.Save(dataDir)
        ' ExEnd:LoadHyphenationDictionaryForLan
        Console.WriteLine(Convert.ToString(vbLf & "Hyphenation dictionary for special language loaded successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
