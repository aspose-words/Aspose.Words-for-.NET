Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection

Imports Aspose.Words
Imports Aspose.Words.Saving

Public Class SaveAsMultipageTiff
    Public Shared Sub Run()
        ' ExStart:SaveAsMultipageTiff
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        ' Open the document.
        Dim doc As New Document(dataDir & "TestFile Multipage TIFF.doc")

        ' ExStart:SaveAsTIFF
        ' Save the document as multipage TIFF.
        doc.Save(dataDir & "TestFile Multipage TIFF_out.tiff")
        ' ExEnd:SaveAsTIFF
        ' ExStart:SaveAsTIFFUsingImageSaveOptions
        ' Create an ImageSaveOptions object to pass to the Save method
        Dim options As New ImageSaveOptions(SaveFormat.Tiff)
        options.PageIndex = 0
        options.PageCount = 2
        options.TiffCompression = TiffCompression.Ccitt4
        options.Resolution = 160

        doc.Save(dataDir & "TestFileWithOptions_out.tiff", options)
        ' ExEnd:SaveAsTIFFUsingImageSaveOptions
        ' ExEnd:SaveAsMultipageTiff
        Console.WriteLine(vbNewLine & "Document saved as multi-page TIFF successfully." & vbNewLine & "File saved at " + dataDir + "TestFileWithOptions Out.tiff")
    End Sub
End Class
