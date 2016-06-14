Imports System.IO
Imports System.Reflection
Imports Aspose.Words.Fonts
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class EmbeddingWindowsStandardFonts
    Public Shared Sub Run()
        ' ExStart:AvoidEmbeddingCoreFonts
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
        Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))
        ' To disable embedding of core fonts and subsuite PDF type 1 fonts set UseCoreFonts to true.
        Dim options As New PdfSaveOptions()
        options.UseCoreFonts = True

        Dim outPath As String = dataDir & Convert.ToString("Rendering.DisableEmbedWindowsFonts_out_.pdf")
        ' The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        doc.Save(outPath)
        ' ExEnd:AvoidEmbeddingCoreFonts
        Console.WriteLine(Convert.ToString(vbLf & "Avoid embedded core fonts setup successfully." & vbLf & "File saved at ") & outPath)
        SkipEmbeddedArialAndTimesRomanFonts(doc, dataDir)
    End Sub
    Private Shared Sub SkipEmbeddedArialAndTimesRomanFonts(doc As Document, dataDir As String)
        ' ExStart:SkipEmbeddedArialAndTimesRomanFonts
        ' To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
        ' To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
        Dim options As New PdfSaveOptions()
        options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll

        dataDir = dataDir & Convert.ToString("Rendering.DisableEmbedWindowsFonts_out_.pdf")
        ' The output PDF will be saved without embedding standard windows fonts.
        doc.Save(dataDir)
        ' ExEnd:SkipEmbeddedArialAndTimesRomanFonts
        Console.WriteLine(Convert.ToString(vbLf & "Embedded arial and times new roman fonts are skipped successfully." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
