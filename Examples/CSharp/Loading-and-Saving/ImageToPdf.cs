using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Words.Drawing;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ImageToPdf
    {
        public static void Run()
        {
            //ExStart:ImageToPdf
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            ConvertImageToPdf(dataDir + "Test.jpg", dataDir + "TestJpg_out_.pdf");
            ConvertImageToPdf(dataDir + "Test.png", dataDir + "TestPng_out_.pdf");
            ConvertImageToPdf(dataDir + "Test.wmf", dataDir + "TestWmf_out_.pdf");
            ConvertImageToPdf(dataDir + "Test.tiff", dataDir + "TestTif_out_.pdf");
            ConvertImageToPdf(dataDir + "Test.gif", dataDir + "TestGif_out_.pdf");
            //ExEnd:ImageToPdf
            Console.WriteLine("\nConverted all images to PDF successfully.");
        }

        /// <summary>
        /// Converts an image to PDF using Aspose.Words for .NET.
        /// </summary>
        /// <param name="inputFileName">File name of input image file.</param>
        /// <param name="outputFileName">Output PDF file name.</param>
        
        public static void ConvertImageToPdf(string inputFileName, string outputFileName)
        {
            Console.WriteLine("Converting " + inputFileName + " to PDF ....");
            //ExStart:ConvertImageToPdf
            // Create Document and DocumentBuilder. 
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
            //ExEnd:ConvertImageToPdf

        }
    }
}
