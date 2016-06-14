using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Drawing;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Images
{
    class CompressImages
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithImages();
            string fileName = "Test.docx";
            string srcFileName = dataDir + fileName;

            Console.WriteLine("Loading {0}. Size {1}.", srcFileName, GetFileSize(srcFileName));
            Document doc = new Document(srcFileName);

            // 220ppi Print - said to be excellent on most printers and screens.
            // 150ppi Screen - said to be good for web pages and projectors.
            // 96ppi Email - said to be good for minimal document size and sharing.
            const int desiredPpi = 150;

            // In .NET this seems to be a good compression / quality setting.
            const int jpegQuality = 90;

            // Resample images to desired ppi and save.
            int count = Resampler.Resample(doc, desiredPpi, jpegQuality);

            Console.WriteLine("Resampled {0} images.", count);

            if (count != 1)
                Console.WriteLine("We expected to have only 1 image resampled in this test document!");

            string dstFileName = dataDir +  RunExamples.GetOutputFilePath(fileName);
            doc.Save(dstFileName);
            Console.WriteLine("Saving {0}. Size {1}.", dstFileName, GetFileSize(dstFileName));

            // Verify that the first image was compressed by checking the new Ppi.
            doc = new Document(dstFileName);
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            double imagePpi = shape.ImageData.ImageSize.WidthPixels / ConvertUtil.PointToInch(shape.SizeInPoints.Width);

            Debug.Assert(imagePpi < 150, "Image was not resampled successfully.");

            Console.WriteLine("\nCompressed images successfully.\nFile saved at " + dstFileName);
        }
        public static int GetFileSize(string fileName)
        {
            using (Stream stream = File.OpenRead(fileName))
                return (int)stream.Length;
        }
    }

    public class Resampler
    {
        /// <summary>
        /// Resamples all images in the document that are greater than the specified PPI (pixels per inch) to the specified PPI
        /// and converts them to JPEG with the specified quality setting.
        /// </summary>
        /// <param name="doc">The document to process.</param>
        /// <param name="desiredPpi">Desired pixels per inch. 220 high quality. 150 screen quality. 96 email quality.</param>
        /// <param name="jpegQuality">0 - 100% JPEG quality.</param>
        /// <returns></returns>
        public static int Resample(Document doc, int desiredPpi, int jpegQuality)
        {
            int count = 0;

            // Convert VML shapes.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                // It is important to use this method to correctly get the picture shape size in points even if the picture is inside a group shape.
                SizeF shapeSizeInPoints = shape.SizeInPoints;

                if (ResampleCore(shape.ImageData, shapeSizeInPoints, desiredPpi, jpegQuality))
                    count++;
            }

            return count;
        }

        /// <summary>
        /// Resamples one VML or DrawingML image
        /// </summary>
        private static bool ResampleCore(ImageData imageData, SizeF shapeSizeInPoints, int ppi, int jpegQuality)
        {
            // The are actually several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
            if (imageData == null)
                return false;

            // An image can be stored in the shape or linked from somewhere else. Let's skip images that do not store bytes in the shape.
            byte[] originalBytes = imageData.ImageBytes;
            if (originalBytes == null)
                return false;

            // Ignore metafiles, they are vector drawings and we don't want to resample them.
            ImageType imageType = imageData.ImageType;
            if (imageType.Equals(ImageType.Wmf) || imageType.Equals(ImageType.Emf))
                return false;

            try
            {
                double shapeWidthInches = ConvertUtil.PointToInch(shapeSizeInPoints.Width);
                double shapeHeightInches = ConvertUtil.PointToInch(shapeSizeInPoints.Height);

                // Calculate the current PPI of the image.
                ImageSize imageSize = imageData.ImageSize;
                double currentPpiX = imageSize.WidthPixels / shapeWidthInches;
                double currentPpiY = imageSize.HeightPixels / shapeHeightInches;

                Console.Write("Image PpiX:{0}, PpiY:{1}. ", (int)currentPpiX, (int)currentPpiY);

                // Let's resample only if the current PPI is higher than the requested PPI (e.g. we have extra data we can get rid of).
                if ((currentPpiX <= ppi) || (currentPpiY <= ppi))
                {
                    Console.WriteLine("Skipping.");
                    return false;
                }

                using (Image srcImage = imageData.ToImage())
                {
                    // Create a new image of such size that it will hold only the pixels required by the desired ppi.
                    int dstWidthPixels = (int)(shapeWidthInches * ppi);
                    int dstHeightPixels = (int)(shapeHeightInches * ppi);
                    using (Bitmap dstImage = new Bitmap(dstWidthPixels, dstHeightPixels))
                    {
                        // Drawing the source image to the new image scales it to the new size.
                        using (Graphics gr = Graphics.FromImage(dstImage))
                        {
                            gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            gr.DrawImage(srcImage, 0, 0, dstWidthPixels, dstHeightPixels);
                        }

                        // Create JPEG encoder parameters with the quality setting.
                        ImageCodecInfo encoderInfo = GetEncoderInfo(ImageFormat.Jpeg);
                        EncoderParameters encoderParams = new EncoderParameters();
                        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, jpegQuality);

                        // Save the image as JPEG to a memory stream.
                        MemoryStream dstStream = new MemoryStream();
                        dstImage.Save(dstStream, encoderInfo, encoderParams);

                        // If the image saved as JPEG is smaller than the original, store it in the shape.
                        Console.WriteLine("Original size {0}, new size {1}.", originalBytes.Length, dstStream.Length);
                        if (dstStream.Length < originalBytes.Length)
                        {
                            dstStream.Position = 0;
                            imageData.SetImage(dstStream);
                            return true;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // Catch an exception, log an error and continue if cannot process one of the images for whatever reason.
                Console.WriteLine("Error processing an image, ignoring. " + e.Message);
            }

            return false;
        }

        /// <summary>
        /// Gets the codec info for the specified image format. Throws if cannot find.
        /// </summary>
        private static ImageCodecInfo GetEncoderInfo(ImageFormat format)
        {
            ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();

            for (int i = 0; i < encoders.Length; i++)
            {
                if (encoders[i].FormatID == format.Guid)
                    return encoders[i];
            }

            throw new Exception("Cannot find a codec.");
        }
    }
}
