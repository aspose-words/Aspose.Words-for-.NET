

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection

Imports Aspose.Words
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Aspose.Words.Drawing

Public Class ImageToPdf

    Public Shared Sub Run()
        ' ExStart:ImageToPdf
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ConvertImageToPdf(dataDir & Convert.ToString("Test.jpg"), dataDir & Convert.ToString("TestJpg_out.pdf"))
        ConvertImageToPdf(dataDir & Convert.ToString("Test.png"), dataDir & Convert.ToString("TestPng_out.pdf"))
        ConvertImageToPdf(dataDir & Convert.ToString("Test.wmf"), dataDir & Convert.ToString("TestWmf_out.pdf"))
        ConvertImageToPdf(dataDir & Convert.ToString("Test.tiff"), dataDir & Convert.ToString("TestTiff_out.pdf"))
        ConvertImageToPdf(dataDir & Convert.ToString("Test.gif"), dataDir & Convert.ToString("TestGif_out.pdf"))
        ' ExEnd:ImageToPdf

        Console.WriteLine(vbLf & "Converted all images to PDF successfully.")
    End Sub

    ''' <summary>
    ''' Converts an image to PDF using Aspose.Words for .NET.
    ''' </summary>
    ''' <param name="inputFileName">File name of input image file.</param>
    ''' <param name="outputFileName">Output PDF file name.</param>
    Public Shared Sub ConvertImageToPdf(inputFileName As String, outputFileName As String)
        Console.WriteLine((Convert.ToString("Converting ") & inputFileName) + " to PDF ....")
        ' ExStart:ConvertImageToPdf
        ' Create Aspose.Words.Document and DocumentBuilder. 
        ' The builder makes it simple to add content to the document.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Read the image from file, ensure it is disposed.
        Using image__1 As Image = Image.FromFile(inputFileName)
            ' Find which dimension the frames in this image represent. For example 
            ' The frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension". 
            Dim dimension As New FrameDimension(image__1.FrameDimensionsList(0))

            ' Get the number of frames in the image.
            Dim framesCount As Integer = image__1.GetFrameCount(dimension)

            ' Loop through all frames.
            For frameIdx As Integer = 0 To framesCount - 1
                ' Insert a section break before each new page, in case of a multi-frame TIFF.
                If frameIdx <> 0 Then
                    builder.InsertBreak(BreakType.SectionBreakNewPage)
                End If

                ' Select active frame.
                image__1.SelectActiveFrame(dimension, frameIdx)

                ' We want the size of the page to be the same as the size of the image.
                ' Convert pixels to points to size the page to the actual image size.
                Dim ps As PageSetup = builder.PageSetup
                ps.PageWidth = ConvertUtil.PixelToPoint(image__1.Width, image__1.HorizontalResolution)
                ps.PageHeight = ConvertUtil.PixelToPoint(image__1.Height, image__1.VerticalResolution)

                ' Insert the image into the document and position it at the top left corner of the page.
                builder.InsertImage(image__1, RelativeHorizontalPosition.Page, 0, RelativeVerticalPosition.Page, 0, ps.PageWidth, _
                    ps.PageHeight, WrapType.None)
            Next
        End Using

        ' Save the document to PDF.
        doc.Save(outputFileName)
        ' ExEnd:ConvertImageToPdf
    End Sub

End Class
