'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection

Imports Aspose.Words
Imports Aspose.Words.Saving

Namespace ImageColorFiltersExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Open the document.
			Dim doc As New Document(String.Format("{0}{1}", dataDir, "TestFile.docx"))

			SaveColorTIFFwithLZW(doc, dataDir, 0.8f, 0.8f)
			SaveGrayscaleTIFFwithLZW(doc, dataDir, 0.8f, 0.8f)
			SaveBlackWhiteTIFFwithLZW(doc, dataDir, True)
			SaveBlackWhiteTIFFwithCITT4(doc, dataDir, True)
			SaveBlackWhiteTIFFwithRLE(doc, dataDir, True)
		End Sub

		'ExStart
		'ExFor:ImageSaveOptions.ImageContrast
		'ExFor:ImageSaveOptions.ImageBrightness
		'ExId:ImageColorFilters_tiff_lzw_color
		'ExSummary: Applies LZW compression, saves to color TIFF image with specified brightness and contrast.
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
		End Sub
		'ExEnd

		'ExStart
		'ExFor:ImageColorMode
		'ExId:ImageColorFilters_tiff_lzw_grayscale
		'ExSummary: Applies LZW compression, saves to grayscale TIFF image with specified brightness and contrast.
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
		End Sub
		'ExEnd

		'ExStart
		'ExId:ImageColorFilters_tiff_lzw_blackandwhite_sens
		'ExSummary: Applies LZW compression, saves to black & white TIFF image with specified sensitivity to gray color.
		Private Shared Sub SaveBlackWhiteTIFFwithLZW(ByVal doc As Document, ByVal dataDir As String, ByVal highSensitivity As Boolean)
			' Select the TIFF format with 100 dpi.
			Dim imgOpttiff As New ImageSaveOptions(SaveFormat.Tiff)
			imgOpttiff.Resolution = 100

			' Apply black & white filter. Set very high sensitivity to gray color.
			imgOpttiff.TiffCompression = TiffCompression.Lzw
			imgOpttiff.ImageColorMode = ImageColorMode.BlackAndWhite

			' Set brightness and contrast according to sensitivity.
			If highSensitivity Then
				imgOpttiff.ImageBrightness = 0.4f
				imgOpttiff.ImageContrast = 0.3f
			Else
				imgOpttiff.ImageBrightness = 0.9f
				imgOpttiff.ImageContrast = 0.9f
			End If

			' Save multipage TIFF.
			doc.Save(String.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff)
		End Sub
		'ExEnd

		'ExStart
		'ExId:ImageColorFilters_tiff_ccitt4_graysvale_sens
		'ExSummary: Applies CCITT4 compression, saves to black & white TIFF image with specified sensitivity to gray color.
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
				imgOpttiff.ImageBrightness = 0.4f
				imgOpttiff.ImageContrast = 0.3f
			Else
				imgOpttiff.ImageBrightness = 0.9f
				imgOpttiff.ImageContrast = 0.9f
			End If

			' Save multipage TIFF.
			doc.Save(String.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff)
		End Sub
		'ExEnd

		'ExStart
		'ExId:ImageColorFilters_tiff_rle_graysvale_sens
		'ExSummary: Applies RLE compression with specified sensitivity to gray color, saves to black & white TIFF image.
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
				imgOpttiff.ImageBrightness = 0.4f
				imgOpttiff.ImageContrast = 0.3f
			Else
				imgOpttiff.ImageBrightness = 0.9f
				imgOpttiff.ImageContrast = 0.9f
			End If

			' Save multipage TIFF grayscale with low bright and contrast
			doc.Save(String.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff)
		End Sub
		'ExEnd
	End Class
End Namespace