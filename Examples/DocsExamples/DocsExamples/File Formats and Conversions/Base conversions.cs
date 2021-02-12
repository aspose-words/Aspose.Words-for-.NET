using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Email;
using Aspose.Email.Clients.Smtp;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions
{
    public class BaseConversions : DocsExamplesBase
    {
        [Test]
        public void DocToDocx()
        {
            //ExStart:LoadAndSave
            //ExStart:OpenDocument
            Document doc = new Document(MyDir + "Document.doc");
            //ExEnd:OpenDocument

            doc.Save(ArtifactsDir + "BaseConversions.DocToDocx.docx");
            //ExEnd:LoadAndSave
        }

        [Test]
        public void DocxToRtf()
        {
            //ExStart:LoadAndSaveToStream 
            //ExStart:OpeningFromStream
            // Read only access is enough for Aspose.Words to load a document.
            Stream stream = File.OpenRead(MyDir + "Document.docx");

            Document doc = new Document(stream);
            // You can close the stream now, it is no longer needed because the document is in memory.
            stream.Close();
            //ExEnd:OpeningFromStream 

            // ... do something with the document.

            // Convert the document to a different format and save to stream.
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Rtf);

            // Rewind the stream position back to zero so it is ready for the next reader.
            dstStream.Position = 0;
            //ExEnd:LoadAndSaveToStream 
            
            File.WriteAllBytes(ArtifactsDir + "BaseConversions.DocxToRtf.rtf", dstStream.ToArray());
        }

        [Test]
        public void DocxToPdf()
        {
            //ExStart:Doc2Pdf
            Document doc = new Document(MyDir + "Document.docx");

            doc.Save(ArtifactsDir + "BaseConversions.DocxToPdf.pdf");
            //ExEnd:Doc2Pdf
        }

        [Test]
        public void DocxToByte()
        {
            //ExStart:DocxToByte
            Document doc = new Document(MyDir + "Document.docx");

            MemoryStream outStream = new MemoryStream();
            doc.Save(outStream, SaveFormat.Docx);

            byte[] docBytes = outStream.ToArray();
            MemoryStream inStream = new MemoryStream(docBytes);

            Document docFromBytes = new Document(inStream);
            //ExEnd:DocxToByte
        }

        [Test]
        public void DocxToEpub()
        {
            //ExStart:DocxToEpub
            Document doc = new Document(MyDir + "Document.docx");

            doc.Save(ArtifactsDir + "BaseConversions.DocxToEpub.epub");
            //ExEnd:DocxToEpub
        }

        [Test, Ignore("Only for example")]
        public void DocxToMhtmlAndSendingEmail()
        {
            //ExStart:DocxToMhtmlAndSendingEmail
            Document doc = new Document(MyDir + "Document.docx");

            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);

            // Rewind the stream to the beginning so Aspose.Email can read it.
            stream.Position = 0;

            // Create an Aspose.Network MIME email message from the stream.
            MailMessage message = MailMessage.Load(stream, new MhtmlLoadOptions());
            message.From = "your_from@email.com";
            message.To = "your_to@email.com";
            message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

            // Send the message using Aspose.Email.
            SmtpClient client = new SmtpClient();
            client.Host = "your_smtp.com";
            client.Send(message);
            //ExEnd:DocxToMhtmlAndSendingEmail
        }

        [Test]
        public void DocxToMarkdown()
        {
            //ExStart:SaveToMarkdownDocument
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Some text!");

            doc.Save(ArtifactsDir + "BaseConversions.DocxToMarkdown.md");
            //ExEnd:SaveToMarkdownDocument
        }

        [Test]
        public void DocxToTxt()
        {
            //ExStart:DocxToTxt
            Document doc = new Document(MyDir + "Document.docx");

            doc.Save(ArtifactsDir + "BaseConversions.DocxToTxt.txt");
            //ExEnd:DocxToTxt
        }

        [Test]
        public void TxtToDocx()
        {
            //ExStart:TxtToDocx
            // The encoding of the text file is automatically detected.
            Document doc = new Document(MyDir + "English text.txt");

            doc.Save(ArtifactsDir + "BaseConversions.TxtToDocx.docx");
            //ExEnd:TxtToDocx
        }

        [Test]
        public void PdfToJpeg()
        {
            //ExStart:PdfToJpeg
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            doc.Save(ArtifactsDir + "BaseConversions.PdfToJpeg.jpeg");
            //ExEnd:PdfToJpeg
        }

        [Test]
        public void PdfToDocx()
        {
            //ExStart:PdfToDocx
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            doc.Save(ArtifactsDir + "BaseConversions.PdfToDocx.docx");
            //ExEnd:PdfToDocx
        }

#if NET462
        [Test]
        public void ImagesToPdf()
        {
            //ExStart:ImageToPdf
            ConvertImageToPdf(ImagesDir + "Logo.jpg", ArtifactsDir + "BaseConversions.JpgToPdf.pdf");
            ConvertImageToPdf(ImagesDir + "Transparent background logo.png", ArtifactsDir + "BaseConversions.PngToPdf.pdf");
            ConvertImageToPdf(ImagesDir + "Windows MetaFile.wmf", ArtifactsDir + "BaseConversions.WmfToPdf.pdf");
            ConvertImageToPdf(ImagesDir + "Tagged Image File Format.tiff", ArtifactsDir + "BaseConversions.TiffToPdf.pdf");
            ConvertImageToPdf(ImagesDir + "Graphics Interchange Format.gif", ArtifactsDir + "BaseConversions.GifToPdf.pdf");
            //ExEnd:ImageToPdf
        }

        /// <summary>
        /// Converts an image to PDF using Aspose.Words for .NET.
        /// </summary>
        /// <param name="inputFileName">File name of input image file.</param>
        /// <param name="outputFileName">Output PDF file name.</param>
        public void ConvertImageToPdf(string inputFileName, string outputFileName)
        {
            Console.WriteLine("Converting " + inputFileName + " to PDF ....");

            //ExStart:ConvertImageToPdf
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Read the image from file, ensure it is disposed.
            using (Image image = Image.FromFile(inputFileName))
            {
                // Find which dimension the frames in this image represent. For example 
                // the frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension".
                FrameDimension dimension = new FrameDimension(image.FrameDimensionsList[0]);

                int framesCount = image.GetFrameCount(dimension);

                for (int frameIdx = 0; frameIdx < framesCount; frameIdx++)
                {
                    // Insert a section break before each new page, in case of a multi-frame TIFF.
                    if (frameIdx != 0)
                        builder.InsertBreak(BreakType.SectionBreakNewPage);

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

            doc.Save(outputFileName);
            //ExEnd:ConvertImageToPdf
        }
#endif
    }
}