Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection

Imports Aspose.Words
Imports Aspose.Words.Saving

Public Class ImageColorFilters
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()

        ' Open the document.
        Dim doc As New Document(String.Format("{0}{1}", dataDir, "TestFile.docx"))

        SaveColorTIFFwithLZW(doc, dataDir, 0.8F, 0.8F)
        SaveGrayscaleTIFFwithLZW(doc, dataDir, 0.8F, 0.8F)
        SaveBlackWhiteTIFFwithLZW(doc, dataDir, True)
        SaveBlackWhiteTIFFwithCITT4(doc, dataDir, True)
        SaveBlackWhiteTIFFwithRLE(doc, dataDir, True)


    End Sub

    Private Shared Sub SaveColorTIFFwithLZW(ByVal doc As Document, ByVal dataDir As String, ByVal brightness As Single, ByVal contrast As Single)
        ' Select the TIFF format with 100 dpi.
        Dim imgOpttiff As New ImageSaveOptions(SaveFormat.Tiff)
        imgOpttiff.Resolution = 100

        ' Select fullcolor LZW compression.
        imgOpttiff.TiffCompression = TiffCompression.Lzw

        ' Set brightness and contrast.
        imgOpttiff.ImageBrightness = brightness
        imgOpttiff.ImageContrast = contrast

        ' Save multipage color TIFF.
        doc.Save(String.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff)

        Console.WriteLine(vbNewLine & "Document converted to TIFF successfully with Colors." & vbNewLine & "File saved at " + dataDir + "Result Colors.tiff")
    End Sub
    
    Private Shared Sub SaveGrayscaleTIFFwithLZW(ByVal doc As Document, ByVal dataDir As String, ByVal brightness As Single, ByVal contrast As Single)
        ' Select the TIFF format with 100 dpi.
        Dim imgOpttiff As New ImageSaveOptions(SaveFormat.Tiff)
        imgOpttiff.Resolution = 100

        ' Select LZW compression.
        imgOpttiff.TiffCompression = TiffCompression.Lzw

        ' Apply grayscale filter.
        imgOpttiff.ImageColorMode = ImageColorMode.Grayscale

        ' Set brightness and contrast.
        imgOpttiff.ImageBrightness = brightness
        imgOpttiff.ImageContrast = contrast

        ' Save multipage grayscale TIFF.
        doc.Save(String.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff)

        Console.WriteLine(vbNewLine & "Document converted to TIFF successfully with Gray scale." & vbNewLine & "File saved at " + dataDir + "Result Grayscale.tiff")
    End Sub
    
    Private Shared Sub SaveBlackWhiteTIFFwithLZW(ByVal doc As Document, ByVal dataDir As String, ByVal highSensitivity As Boolean)
        ' Select the TIFF format with 100 dpi.
        Dim imgOpttiff As New ImageSaveOptions(SaveFormat.Tiff)
        imgOpttiff.Resolution = 100

        ' Apply black & white filter. Set very high sensitivity to gray color.
        imgOpttiff.TiffCompression = TiffCompression.Lzw
        imgOpttiff.ImageColorMode = ImageColorMode.BlackAndWhite

        ' Set brightness and contrast according to sensitivity.
        If highSensitivity Then
            imgOpttiff.ImageBrightness = 0.4F
            imgOpttiff.ImageContrast = 0.3F
        Else
            imgOpttiff.ImageBrightness = 0.9F
            imgOpttiff.ImageContrast = 0.9F
        End If

        ' Save multipage TIFF.
        doc.Save(String.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff)

        Console.WriteLine(vbNewLine & "Document converted to TIFF successfully with black and white." & vbNewLine & "File saved at " + dataDir + "Result black and white.tiff")
    End Sub
    
    Private Shared Sub SaveBlackWhiteTIFFwithCITT4(ByVal doc As Document, ByVal dataDir As String, ByVal highSensitivity As Boolean)
        ' Select the TIFF format with 100 dpi.
        Dim imgOpttiff As New ImageSaveOptions(SaveFormat.Tiff)
        imgOpttiff.Resolution = 100

        ' Set CCITT4 compression.
        imgOpttiff.TiffCompression = TiffCompression.Ccitt4

        ' Apply grayscale filter.
        imgOpttiff.ImageColorMode = ImageColorMode.Grayscale

        ' Set brightness and contrast according to sensitivity.
        If highSensitivity Then
            imgOpttiff.ImageBrightness = 0.4F
            imgOpttiff.ImageContrast = 0.3F
        Else
            imgOpttiff.ImageBrightness = 0.9F
            imgOpttiff.ImageContrast = 0.9F
        End If

        ' Save multipage TIFF.
        doc.Save(String.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff)

        Console.WriteLine(vbNewLine & "Document converted to TIFF successfully with black and white and Ccitt4 compression." & vbNewLine & "File saved at " + dataDir + "Result Ccitt4.tiff")
    End Sub
    
    Private Shared Sub SaveBlackWhiteTIFFwithRLE(ByVal doc As Document, ByVal dataDir As String, ByVal highSensitivity As Boolean)
        ' Select the TIFF format with 100 dpi.
        Dim imgOpttiff As New ImageSaveOptions(SaveFormat.Tiff)
        imgOpttiff.Resolution = 100

        ' Set RLE compression.
        imgOpttiff.TiffCompression = TiffCompression.Rle

        ' Aply grayscale filter.
        imgOpttiff.ImageColorMode = ImageColorMode.Grayscale

        ' Set brightness and contrast according to sensitivity.
        If highSensitivity Then
            imgOpttiff.ImageBrightness = 0.4F
            imgOpttiff.ImageContrast = 0.3F
        Else
            imgOpttiff.ImageBrightness = 0.9F
            imgOpttiff.ImageContrast = 0.9F
        End If

        ' Save multipage TIFF grayscale with low bright and contrast
        doc.Save(String.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff)

        Console.WriteLine(vbNewLine & "Document converted to TIFF successfully with black and white and Rle compression." & vbNewLine & "File saved at " + dataDir + "Result Rle.tiff")
    End Sub
End Class
