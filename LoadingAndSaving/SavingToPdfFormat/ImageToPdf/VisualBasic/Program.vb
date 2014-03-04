'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection

Imports Aspose.Words
Imports Aspose.Words.Drawing

Namespace ImageToPdfExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			ConvertImageToPdf(dataDir & "Test.jpg", dataDir & "TestJpg Out.pdf")
			ConvertImageToPdf(dataDir & "Test.png", dataDir & "TestPng Out.pdf")
			ConvertImageToPdf(dataDir & "Test.wmf", dataDir & "TestWmf Out.pdf")
			ConvertImageToPdf(dataDir & "Test.tiff", dataDir & "TestTiff Out.pdf")
			ConvertImageToPdf(dataDir & "Test.gif", dataDir & "TestGif Out.pdf")
		End Sub

		''' <summary>
		''' Converts an image to PDF using Aspose.Words for .NET.
		''' </summary>
		''' <param name="inputFileName">File name of input image file.</param>
		''' <param name="outputFileName">Output PDF file name.</param>
		Public Shared Sub ConvertImageToPdf(ByVal inputFileName As String, ByVal outputFileName As String)
			' Create Aspose.Words.Document and DocumentBuilder. 
			' The builder makes it simple to add content to the document.
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			' Read the image from file, ensure it is disposed.
			Using image As Image = Image.FromFile(inputFileName)
				' Find which dimension the frames in this image represent. For example 
				' the frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension". 
				Dim dimension As New FrameDimension(image.FrameDimensionsList(0))

				' Get the number of frames in the image.
				Dim framesCount As Integer = image.GetFrameCount(dimension)

				' Loop through all frames.
				For frameIdx As Integer = 0 To framesCount - 1
					' Insert a section break before each new page, in case of a multi-frame TIFF.
					If frameIdx <> 0 Then
						builder.InsertBreak(BreakType.SectionBreakNewPage)
					End If

					' Select active frame.
					image.SelectActiveFrame(dimension, frameIdx)

					' We want the size of the page to be the same as the size of the image.
					' Convert pixels to points to size the page to the actual image size.
					Dim ps As PageSetup = builder.PageSetup
					ps.PageWidth = ConvertUtil.PixelToPoint(image.Width, image.HorizontalResolution)
					ps.PageHeight = ConvertUtil.PixelToPoint(image.Height, image.VerticalResolution)

					' Insert the image into the document and position it at the top left corner of the page.
					builder.InsertImage(image, RelativeHorizontalPosition.Page, 0, RelativeVerticalPosition.Page, 0, ps.PageWidth, ps.PageHeight, WrapType.None)
				Next frameIdx
			End Using

			' Save the document to PDF.
			doc.Save(outputFileName)
		End Sub
	End Class
End Namespace
'ExEnd