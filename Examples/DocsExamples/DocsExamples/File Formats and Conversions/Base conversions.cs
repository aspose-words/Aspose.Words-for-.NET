using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Email;
using Aspose.Email.Clients.Smtp;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions
{
    public class BaseConversions : DocsExamplesBase
    {
        [Test]
        public void DocToDocx()
        {
            //ExStart:LoadAndSave
            //GistId:7ee438947078cf070c5bc36a4e45a18c
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
            //GistId:7ee438947078cf070c5bc36a4e45a18c
            //ExStart:OpenFromStream
            //GistId:1d626c7186a318d22d022dc96dd91d55
            // Read only access is enough for Aspose.Words to load a document.
            Document doc;
            using (Stream stream = File.OpenRead(MyDir + "Document.docx"))
                doc = new Document(stream);
            //ExEnd:OpenFromStream

            // ... do something with the document.

            // Convert the document to a different format and save to stream.
            using (MemoryStream dstStream = new MemoryStream())
            {
                doc.Save(dstStream, SaveFormat.Rtf);
                // Rewind the stream position back to zero so it is ready for the next reader.
                dstStream.Position = 0;

                File.WriteAllBytes(ArtifactsDir + "BaseConversions.DocxToRtf.rtf", dstStream.ToArray());
            }
            //ExEnd:LoadAndSaveToStream
        }

        [Test]
        public void DocxToPdf()
        {
            //ExStart:DocxToPdf
            //GistId:a53bdaad548845275c1b9556ee21ae65
            Document doc = new Document(MyDir + "Document.docx");

            doc.Save(ArtifactsDir + "BaseConversions.DocxToPdf.pdf");
            //ExEnd:DocxToPdf
        }

        [Test]
        public void DocxToByte()
        {
            //ExStart:DocxToByte
            //GistId:f8a622f8bc1cf3c2fa8a7a9be359faa2
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

        [Test]
        public void DocxToHtml()
        {
            //ExStart:DocxToHtml
            //GistId:c0df00d37081f41a7683339fd7ef66c1
            Document doc = new Document(MyDir + "Document.docx");

            doc.Save(ArtifactsDir + "BaseConversions.DocxToHtml.html");
            //ExEnd:DocxToHtml
        }

        [Test, Ignore("Only for example")]
        public void DocxToMhtml()
        {
            //ExStart:DocxToMhtml
            //GistId:537e7d4e2ddd23fa701dc4bf315064b9
            Document doc = new Document(MyDir + "Document.docx");

            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);

            // Rewind the stream to the beginning so Aspose.Email can read it.
            stream.Position = 0;

            // Create an Aspose.Email MIME email message from the stream.
            MailMessage message = MailMessage.Load(stream, new MhtmlLoadOptions());
            message.From = "your_from@email.com";
            message.To = "your_to@email.com";
            message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

            // Send the message using Aspose.Email.
            SmtpClient client = new SmtpClient();
            client.Host = "your_smtp.com";
            client.Send(message);
            //ExEnd:DocxToMhtml
        }

        [Test]
        public void DocxToMarkdown()
        {
            //ExStart:DocxToMarkdown
            //GistId:51b4cb9c451832f23527892e19c7bca6
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Some text!");

            doc.Save(ArtifactsDir + "BaseConversions.DocxToMarkdown.md");
            //ExEnd:DocxToMarkdown
        }

        [Test]
        public void DocxToTxt()
        {
            //ExStart:DocxToTxt
            //GistId:1f94e59ea4838ffac2f0edf921f67060
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "BaseConversions.DocxToTxt.txt");
            //ExEnd:DocxToTxt
        }

        [Test]
        public void DocxToXlsx()
        {
            //ExStart:DocxToXlsx
            //GistId:f5a08835e924510d3809e41c3b8b81a2
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "BaseConversions.DocxToXlsx.xlsx");
            //ExEnd:DocxToXlsx
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
            //GistId:ebbb90d74ef57db456685052a18f8e86
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            doc.Save(ArtifactsDir + "BaseConversions.PdfToJpeg.jpeg");
            //ExEnd:PdfToJpeg
        }

        [Test]
        public void PdfToDocx()
        {
            //ExStart:PdfToDocx
            //GistId:a0d52b62c1643faa76a465a41537edfc
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            doc.Save(ArtifactsDir + "BaseConversions.PdfToDocx.docx");
            //ExEnd:PdfToDocx
        }

        [Test]
        public void PdfToXlsx()
        {
            //ExStart:PdfToXlsx
            //GistId:a50652f28531278511605e0fd778bbdf
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            doc.Save(ArtifactsDir + "BaseConversions.PdfToXlsx.xlsx");
            //ExEnd:PdfToXlsx
        }

        [Test]
        public void FindReplaceXlsx()
        {
            //ExStart:FindReplaceXlsx
            //GistId:a50652f28531278511605e0fd778bbdf
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Ruby bought a ruby necklace.");

            // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
            FindReplaceOptions options = new FindReplaceOptions();

            // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
            // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
            options.MatchCase = true;

            doc.Range.Replace("Ruby", "Jade", options);

            doc.Save(ArtifactsDir + "BaseConversions.FindReplaceXlsx.xlsx");
            //ExEnd:FindReplaceXlsx
        }

        [Test]
        public void CompressXlsx()
        {
            //ExStart:CompressXlsx
            //GistId:a50652f28531278511605e0fd778bbdf
            Document doc = new Document(MyDir + "Document.docx");

            XlsxSaveOptions saveOptions = new XlsxSaveOptions();
            saveOptions.CompressionLevel = CompressionLevel.Maximum;

            doc.Save(ArtifactsDir + "BaseConversions.CompressXlsx.xlsx", saveOptions);
            //ExEnd:CompressXlsx
        }

#if NET48 || JAVA
        [Test]
        public void ImagesToPdf()
        {
            //ExStart:ImageToPdf
            //GistId:a53bdaad548845275c1b9556ee21ae65
            ConvertImageToPdf(ImagesDir + "Logo.jpg", ArtifactsDir + "BaseConversions.JpgToPdf.pdf");
            ConvertImageToPdf(ImagesDir + "Transparent background logo.png", ArtifactsDir + "BaseConversions.PngToPdf.pdf");
            ConvertImageToPdf(ImagesDir + "Windows MetaFile.wmf", ArtifactsDir + "BaseConversions.WmfToPdf.pdf");
            ConvertImageToPdf(ImagesDir + "Tagged Image File Format.tiff", ArtifactsDir + "BaseConversions.TiffToPdf.pdf");
            ConvertImageToPdf(ImagesDir + "Graphics Interchange Format.gif", ArtifactsDir + "BaseConversions.GifToPdf.pdf");
            //ExEnd:ImageToPdf
        }

        //ExStart:ConvertImageToPdf
        //GistId:a53bdaad548845275c1b9556ee21ae65
        /// <summary>
        /// Converts an image to PDF using Aspose.Words for .NET.
        /// </summary>
        /// <param name="inputFileName">File name of input image file.</param>
        /// <param name="outputFileName">Output PDF file name.</param>
        public void ConvertImageToPdf(string inputFileName, string outputFileName)
        {
            Console.WriteLine("Converting " + inputFileName + " to PDF ....");

            
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
        }
        //ExEnd:ConvertImageToPdf
#endif
    }
}