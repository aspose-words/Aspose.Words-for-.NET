// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// Create Aspose.Words.Document and DocumentBuilder. 
// The builder makes it simple to add content to the document.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Read the image from file, ensure it is disposed.
using (Image image = Image.FromFile(inputFileName))
{
    // Find which dimension the frames in this image represent. For example 
    // the frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension". 
    FrameDimension dimension = new FrameDimension(image.FrameDimensionsList[0]);

    // Get the number of frames in the image.
    int framesCount = image.GetFrameCount(dimension);

    // Loop through all frames.
    for (int frameIdx = 0; frameIdx < framesCount; frameIdx++)
    {
        // Insert a section break before each new page, in case of a multi-frame TIFF.
        if (frameIdx != 0)
            builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Select active frame.
        image.SelectActiveFrame(dimension, frameIdx);

        // We want the size of the page to be the same as the size of the image.
        // Convert pixels to points to size the page to the actual image size.
        PageSetup ps = builder.PageSetup;
        ps.PageWidth = ConvertUtil.PixelToPoint(image.Width, image.HorizontalResolution);
        ps.PageHeight = ConvertUtil.PixelToPoint(image.Height, image.VerticalResolution);

        // Insert the image into the document and position it at the top left corner of the page.
        builder.InsertImage(
            image,
            RelativeHorizontalPosition.Page,
            0,
            RelativeVerticalPosition.Page,
            0,
            ps.PageWidth,
            ps.PageHeight,
            WrapType.None);
    }
}

// Save the document to PDF.
doc.Save(outputFileName);
