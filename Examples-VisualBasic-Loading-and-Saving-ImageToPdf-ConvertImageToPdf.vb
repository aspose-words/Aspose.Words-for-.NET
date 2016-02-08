' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Create Aspose.Words.Document and DocumentBuilder. 
' The builder makes it simple to add content to the document.
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

' Read the image from file, ensure it is disposed.
Using image__1 As Image = Image.FromFile(inputFileName)
    ' Find which dimension the frames in this image represent. For example 
    ' the frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension". 
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
