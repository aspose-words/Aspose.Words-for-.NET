#if NET462
using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Layout;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements
{
    internal class WorkingWithImages : DocsExamplesBase
    {
        [Test]
        public void AddImageToEachPage()
        {
            Document doc = new Document(MyDir + "Document.docx");

            // Create and attach collector before the document before page layout is built.
            LayoutCollector layoutCollector = new LayoutCollector(doc);

            // Images in a document are added to paragraphs to add an image to every page we need
            // to find at any paragraph belonging to each page.
            IEnumerator enumerator = doc.SelectNodes("// Body/Paragraph").GetEnumerator();

            for (int page = 1; page <= doc.PageCount; page++)
            {
                while (enumerator.MoveNext())
                {
                    // Check if the current paragraph belongs to the target page.
                    Paragraph paragraph = (Paragraph) enumerator.Current;
                    if (layoutCollector.GetStartPageIndex(paragraph) == page)
                    {
                        AddImageToPage(paragraph, page, ImagesDir);
                        break;
                    }
                }
            }

            // If we need to save the document as a PDF or image, call UpdatePageLayout() method.
            doc.UpdatePageLayout();

            doc.Save(ArtifactsDir + "WorkingWithImages.AddImageToEachPage.docx");
        }

        /// <summary>
        /// Adds an image to a page using the supplied paragraph.
        /// </summary>
        /// <param name="para">The paragraph to an an image to.</param>
        /// <param name="page">The page number the paragraph appears on.</param>
        public void AddImageToPage(Paragraph para, int page, string imagesDir)
        {
            Document doc = (Document) para.Document;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(para);

            // Insert a logo to the top left of the page to place it in front of all other text.
            builder.InsertImage(ImagesDir + "Transparent background logo.png", RelativeHorizontalPosition.Page, 60,
                RelativeVerticalPosition.Page, 60, -1, -1, WrapType.None);

            // Insert a textbox next to the image which contains some text consisting of the page number.
            Shape textBox = new Shape(doc, ShapeType.TextBox);

            // We want a floating shape relative to the page.
            textBox.WrapType = WrapType.None;
            textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            textBox.Height = 30;
            textBox.Width = 200;
            textBox.Left = 150;
            textBox.Top = 80;

            textBox.AppendChild(new Paragraph(doc));
            builder.InsertNode(textBox);
            builder.MoveTo(textBox.FirstChild);
            builder.Writeln("This is a custom note for page " + page);
        }

        [Test]
        public void InsertBarcodeImage()
        {
            //ExStart:InsertBarcodeImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The number of pages the document should have
            const int numPages = 4;
            // The document starts with one section, insert the barcode into this existing section
            InsertBarcodeIntoFooter(builder, doc.FirstSection, HeaderFooterType.FooterPrimary);

            for (int i = 1; i < numPages; i++)
            {
                // Clone the first section and add it into the end of the document
                Section cloneSection = (Section) doc.FirstSection.Clone(false);
                cloneSection.PageSetup.SectionStart = SectionStart.NewPage;
                doc.AppendChild(cloneSection);

                // Insert the barcode and other information into the footer of the section
                InsertBarcodeIntoFooter(builder, cloneSection, HeaderFooterType.FooterPrimary);
            }

            // Save the document as a PDF to disk
            // You can also save this directly to a stream
            doc.Save(ArtifactsDir + "InsertBarcodeImage.docx");
            //ExEnd:InsertBarcodeImage
        }

        //ExStart:InsertBarcodeIntoFooter
        private void InsertBarcodeIntoFooter(DocumentBuilder builder, Section section,
            HeaderFooterType footerType)
        {
            // Move to the footer type in the specific section.
            builder.MoveToSection(section.Document.IndexOf(section));
            builder.MoveToHeaderFooter(footerType);

            // Insert the barcode, then move to the next line and insert the ID along with the page number.
            // Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.
            builder.InsertImage(Image.FromFile(ImagesDir + "Barcode.png"));
            builder.Writeln();
            builder.Write("1234567890");
            builder.InsertField("PAGE");

            // Create a right-aligned tab at the right margin.
            double tabPos = section.PageSetup.PageWidth - section.PageSetup.RightMargin - section.PageSetup.LeftMargin;
            builder.CurrentParagraph.ParagraphFormat.TabStops.Add(new TabStop(tabPos, TabAlignment.Right,
                TabLeader.None));

            // Move to the right-hand side of the page and insert the page and page total.
            builder.Write(ControlChar.Tab);
            builder.InsertField("PAGE");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES");
        }
        //ExEnd:InsertBarcodeIntoFooter

        [Test]
        public void CompressImages()
        {
            Document doc = new Document(MyDir + "Images.docx");

            // 220ppi Print - said to be excellent on most printers and screens.
            // 150ppi Screen - said to be good for web pages and projectors.
            // 96ppi Email - said to be good for minimal document size and sharing.
            const int desiredPpi = 150;

            // In .NET this seems to be a good compression/quality setting.
            const int jpegQuality = 90;

            // Resample images to the desired PPI and save.
            int count = Resampler.Resample(doc, desiredPpi, jpegQuality);

            Console.WriteLine("Resampled {0} images.", count);

            if (count != 1)
                Console.WriteLine("We expected to have only 1 image resampled in this test document!");

            doc.Save(ArtifactsDir + "CompressImages.docx");

            // Verify that the first image was compressed by checking the new PPI.
            doc = new Document(ArtifactsDir + "CompressImages.docx");

            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            double imagePpi = shape.ImageData.ImageSize.WidthPixels / ConvertUtil.PointToInch(shape.SizeInPoints.Width);

            Debug.Assert(imagePpi < 150, "Image was not resampled successfully.");
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

            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                // It is important to use this method to get the picture shape size in points correctly,
                // even if it is inside a group shape.
                SizeF shapeSizeInPoints = shape.SizeInPoints;

                if (ResampleCore(shape.ImageData, shapeSizeInPoints, desiredPpi, jpegQuality))
                    count++;
            }

            return count;
        }

        /// <summary>
        /// Resamples one VML or DrawingML image.
        /// </summary>
        private static bool ResampleCore(ImageData imageData, SizeF shapeSizeInPoints, int ppi, int jpegQuality)
        {
            // The are several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
            if (imageData == null)
                return false;

            // An image can be stored in shape or linked somewhere else, let's skip images that do not store bytes in shape.
            byte[] originalBytes = imageData.ImageBytes;
            if (originalBytes == null)
                return false;

            // Ignore metafiles, they are vector drawings, and we don't want to resample them.
            ImageType imageType = imageData.ImageType;
            if (imageType == ImageType.Wmf || imageType == ImageType.Emf)
                return false;

            try
            {
                double shapeWidthInches = ConvertUtil.PointToInch(shapeSizeInPoints.Width);
                double shapeHeightInches = ConvertUtil.PointToInch(shapeSizeInPoints.Height);

                // Calculate the current PPI of the image.
                ImageSize imageSize = imageData.ImageSize;
                double currentPpiX = imageSize.WidthPixels / shapeWidthInches;
                double currentPpiY = imageSize.HeightPixels / shapeHeightInches;

                Console.Write("Image PpiX:{0}, PpiY:{1}. ", (int) currentPpiX, (int) currentPpiY);

                // Let's resample only if the current PPI is higher than the requested PPI (e.g., we have extra data we can get rid of).
                if (currentPpiX <= ppi || currentPpiY <= ppi)
                {
                    Console.WriteLine("Skipping.");
                    return false;
                }

                using (Image srcImage = imageData.ToImage())
                {
                    // Create a new image of such size that it will hold only the pixels required by the desired PPI.
                    int dstWidthPixels = (int) (shapeWidthInches * ppi);
                    int dstHeightPixels = (int) (shapeHeightInches * ppi);
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

                        // If the image saved as JPEG is smaller than the original, store it in shape.
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
                // Catch an exception, log an error, and continue to process one of the images for whatever reason.
                Console.WriteLine("Error processing an image, ignoring. " + e.Message);
            }

            return false;
        }

        /// <summary>
        /// Gets the codec info for the specified image format.
        /// Throws if cannot find.
        /// </summary>
        private static ImageCodecInfo GetEncoderInfo(ImageFormat format)
        {
            ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();

            foreach (ImageCodecInfo codecInfo in encoders)
            {
                if (codecInfo.FormatID == format.Guid)
                    return codecInfo;
            }

            throw new Exception("Cannot find a codec.");
        }
    }
}
#endif
