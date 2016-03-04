Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Layout
Imports Aspose.Words.Rendering
Public Class HyphenateWordsOfLanguages
    Public Shared Sub Run()
        ' ExStart:HyphenateWordsOfLanguages
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        ' Load the documents which store the shapes we want to render.
        Dim doc As New Document(dataDir & Convert.ToString("TestFile RenderShape.doc"))
        Hyphenation.RegisterDictionary("en-US", dataDir & Convert.ToString("hyph_en_US.dic"))
        Hyphenation.RegisterDictionary("de-CH", dataDir & Convert.ToString("hyph_de_CH.dic"))

        dataDir = dataDir & Convert.ToString("HyphenateWordsOfLanguages_out_.pdf")
        doc.Save(dataDir)
        'ExEnd:HyphenateWordsOfLanguages
        Console.WriteLine(Convert.ToString(vbLf & "Words of special languages hyphenate successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
