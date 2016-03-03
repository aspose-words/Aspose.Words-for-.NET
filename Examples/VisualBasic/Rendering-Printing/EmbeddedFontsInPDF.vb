Imports System.IO
Imports System.Reflection
Imports Aspose.Words.Fonts
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class EmbeddedFontsInPDF
    Public Shared Sub Run()
        ' ExStart:EmbeddAllFonts
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
        Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))
        ' Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
        ' each time a document is rendered.
        Dim options As New PdfSaveOptions()
        options.EmbedFullFonts = True

        Dim outPath As String = dataDir & Convert.ToString("Rendering.EmbedFullFonts_out_.pdf")
        ' The output PDF will be embedded with all fonts found in the document.
        doc.Save(outPath, options)
        ' ExEnd:EmbeddAllFonts
        Console.WriteLine(Convert.ToString(vbLf & "All Fonts embedded successfully." & vbLf & "File saved at ") & outPath)
        EmbeddSubsetFonts(doc, dataDir)
    End Sub
    Private Shared Sub EmbeddSubsetFonts(doc As Document, dataDir As String)
        ' ExStart:EmbeddSubsetFonts
        ' To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
        Dim options As New PdfSaveOptions()
        options.EmbedFullFonts = False
        dataDir = dataDir & Convert.ToString("Rendering.SubsetFonts_out_.pdf")
        ' The output PDF will contain subsets of the fonts in the document. Only the glyphs used
        ' in the document are included in the PDF fonts.
        doc.Save(dataDir, options)
        ' ExEnd:EmbeddSubsetFonts
        Console.WriteLine(Convert.ToString(vbLf & "Subset Fonts embedded successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
