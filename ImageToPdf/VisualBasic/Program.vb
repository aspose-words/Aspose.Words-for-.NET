'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
'ExStart
'ExId:ImageToPdf
'ExSummary:Converts an image into a PDF document.

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection

Imports Aspose.Words
Imports Aspose.Words.Drawing

Namespace ImageToPdf
	Friend Class Program
		Public Shared Sub Main(ByVal args() As String)
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath

			ConvertImageToPdf(dataDir & "Test.jpg", dataDir & "TestJpg Out.pdf")
			ConvertImageToPdf(dataDir & "Test.png", dataDir & "TestPng Out.pdf")
			ConvertImageToPdf(dataDir & "Test.wmf", dataDir & "TestWmf Out.pdf")
			ConvertImageToPdf(dataDir & "Test.tiff", dataDir & "TestTiff Out.pdf")
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
				' Get the number of frames in the image.
				Dim framesCount As Integer = image.GetFrameCount(FrameDimension.Page)

				' Loop through all frames.
				For frameIdx As Integer = 0 To framesCount - 1
					' Insert a section break before each new page, in case of a multi-frame TIFF.
					If frameIdx <> 0 Then
						builder.InsertBreak(BreakType.SectionBreakNewPage)
					End If

					' Select active frame.
					image.SelectActiveFrame(FrameDimension.Page, frameIdx)

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