// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExFile : ApiExampleBase
    {
        [Test]
        public void CatchFileCorruptedException()
        {
            //ExStart
            //ExFor:FileCorruptedException
            //ExSummary:Shows how to catch a FileCorruptedException.
            try
            {
                // If we get an "Unreadable content" error message when trying to open a document using Microsoft Word,
                // chances are that we will get an exception thrown when trying to load that document using Aspose.Words.
                Document doc = new Document(MyDir + "Corrupted document.docx");
            }
            catch (FileCorruptedException e)
            {
                Console.WriteLine(e.Message);
            }
            //ExEnd
        }

        [Test]
        public void DetectEncoding()
        {
            //ExStart
            //ExFor:FileFormatInfo.Encoding
            //ExFor:FileFormatUtil
            //ExSummary:Shows how to detect encoding in an html file.
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Document.html");

            Assert.AreEqual(LoadFormat.Html, info.LoadFormat);

            // The Encoding property is used only when we create a FileFormatInfo object for an html document.
            Assert.AreEqual("Western European (Windows)", info.Encoding.EncodingName);
            Assert.AreEqual(1252, info.Encoding.CodePage);
            //ExEnd

            info = FileFormatUtil.DetectFileFormat(MyDir + "Document.docx");

            Assert.AreEqual(LoadFormat.Docx, info.LoadFormat);
            Assert.IsNull(info.Encoding);
        }

        [Test]
        public void FileFormatToString()
        {
            //ExStart
            //ExFor:FileFormatUtil.ContentTypeToLoadFormat(String)
            //ExFor:FileFormatUtil.ContentTypeToSaveFormat(String)
            //ExSummary:Shows how to find the corresponding Aspose load/save format from each media type string.
            // The ContentTypeToSaveFormat/ContentTypeToLoadFormat methods only accept official IANA media type names, also known as MIME types. 
            // All valid media types are listed here: https://www.iana.org/assignments/media-types/media-types.xhtml.

            // Trying to associate a SaveFormat with a partial media type string will not work.
            Assert.Throws<ArgumentException>(() => FileFormatUtil.ContentTypeToSaveFormat("jpeg"));

            // If Aspose.Words does not have a corresponding save/load format for a content type, an exception will also be thrown.
            Assert.Throws<ArgumentException>(() => FileFormatUtil.ContentTypeToSaveFormat("application/zip"));

            // Files of the types listed below can be saved, but not loaded using Aspose.Words.
            Assert.Throws<ArgumentException>(() => FileFormatUtil.ContentTypeToLoadFormat("image/jpeg"));

            Assert.AreEqual(SaveFormat.Jpeg, FileFormatUtil.ContentTypeToSaveFormat("image/jpeg"));
            Assert.AreEqual(SaveFormat.Png, FileFormatUtil.ContentTypeToSaveFormat("image/png"));
            Assert.AreEqual(SaveFormat.Tiff, FileFormatUtil.ContentTypeToSaveFormat("image/tiff"));
            Assert.AreEqual(SaveFormat.Gif, FileFormatUtil.ContentTypeToSaveFormat("image/gif"));
            Assert.AreEqual(SaveFormat.Emf, FileFormatUtil.ContentTypeToSaveFormat("image/x-emf"));
            Assert.AreEqual(SaveFormat.Xps, FileFormatUtil.ContentTypeToSaveFormat("application/vnd.ms-xpsdocument"));
            Assert.AreEqual(SaveFormat.Pdf, FileFormatUtil.ContentTypeToSaveFormat("application/pdf"));
            Assert.AreEqual(SaveFormat.Svg, FileFormatUtil.ContentTypeToSaveFormat("image/svg+xml"));
            Assert.AreEqual(SaveFormat.Epub, FileFormatUtil.ContentTypeToSaveFormat("application/epub+zip"));

            // For file types that can be saved and loaded, we can match a media type to both a load format and a save format.
            Assert.AreEqual(LoadFormat.Doc, FileFormatUtil.ContentTypeToLoadFormat("application/msword"));
            Assert.AreEqual(SaveFormat.Doc, FileFormatUtil.ContentTypeToSaveFormat("application/msword"));

            Assert.AreEqual(LoadFormat.Docx,
                FileFormatUtil.ContentTypeToLoadFormat(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
            Assert.AreEqual(SaveFormat.Docx,
                FileFormatUtil.ContentTypeToSaveFormat(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));

            Assert.AreEqual(LoadFormat.Text, FileFormatUtil.ContentTypeToLoadFormat("text/plain"));
            Assert.AreEqual(SaveFormat.Text, FileFormatUtil.ContentTypeToSaveFormat("text/plain"));

            Assert.AreEqual(LoadFormat.Rtf, FileFormatUtil.ContentTypeToLoadFormat("application/rtf"));
            Assert.AreEqual(SaveFormat.Rtf, FileFormatUtil.ContentTypeToSaveFormat("application/rtf"));

            Assert.AreEqual(LoadFormat.Html, FileFormatUtil.ContentTypeToLoadFormat("text/html"));
            Assert.AreEqual(SaveFormat.Html, FileFormatUtil.ContentTypeToSaveFormat("text/html"));

            Assert.AreEqual(LoadFormat.Mhtml, FileFormatUtil.ContentTypeToLoadFormat("multipart/related"));
            Assert.AreEqual(SaveFormat.Mhtml, FileFormatUtil.ContentTypeToSaveFormat("multipart/related"));
            //ExEnd
        }

        [Test]
        public void DetectDocumentEncryption()
        {
            //ExStart
            //ExFor:FileFormatUtil.DetectFileFormat(String)
            //ExFor:FileFormatInfo
            //ExFor:FileFormatInfo.LoadFormat
            //ExFor:FileFormatInfo.IsEncrypted
            //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and encryption.
            Document doc = new Document();
            
            // Configure a SaveOptions object to encrypt the document
            // with a password when we save it, and then save the document.
            OdtSaveOptions saveOptions = new OdtSaveOptions(SaveFormat.Odt);
            saveOptions.Password = "MyPassword";

            doc.Save(ArtifactsDir + "File.DetectDocumentEncryption.odt", saveOptions);

            // Verify the file type of our document, and its encryption status.
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(ArtifactsDir + "File.DetectDocumentEncryption.odt");

            Assert.AreEqual(".odt", FileFormatUtil.LoadFormatToExtension(info.LoadFormat));
            Assert.True(info.IsEncrypted);
            //ExEnd
        }

        [Test]
        public void DetectDigitalSignatures()
        {
            //ExStart
            //ExFor:FileFormatUtil.DetectFileFormat(String)
            //ExFor:FileFormatInfo
            //ExFor:FileFormatInfo.LoadFormat
            //ExFor:FileFormatInfo.HasDigitalSignature
            //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and presence of digital signatures.
            // Use a FileFormatInfo instance to verify that a document is not digitally signed.
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Document.docx");

            Assert.AreEqual(".docx", FileFormatUtil.LoadFormatToExtension(info.LoadFormat));
            Assert.False(info.HasDigitalSignature);

            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);
            DigitalSignatureUtil.Sign(MyDir + "Document.docx", ArtifactsDir + "File.DetectDigitalSignatures.docx",
                certificateHolder, new SignOptions() { SignTime = DateTime.Now });

            // Use a new FileFormatInstance to confirm that it is signed.
            info = FileFormatUtil.DetectFileFormat(ArtifactsDir + "File.DetectDigitalSignatures.docx");

            Assert.True(info.HasDigitalSignature);

            // We can load and access the signatures of a signed document in a collection like this.
            Assert.AreEqual(1, DigitalSignatureUtil.LoadSignatures(ArtifactsDir + "File.DetectDigitalSignatures.docx").Count);
            //ExEnd
        }

        [Test]
        public void SaveToDetectedFileFormat()
        {
            //ExStart
            //ExFor:FileFormatUtil.DetectFileFormat(Stream)
            //ExFor:FileFormatUtil.LoadFormatToExtension(LoadFormat)
            //ExFor:FileFormatUtil.ExtensionToSaveFormat(String)
            //ExFor:FileFormatUtil.SaveFormatToExtension(SaveFormat)
            //ExFor:FileFormatUtil.LoadFormatToSaveFormat(LoadFormat)
            //ExFor:Document.OriginalFileName
            //ExFor:FileFormatInfo.LoadFormat
            //ExFor:LoadFormat
            //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document.
            // Load a document from a file that is missing a file extension, and then detect its file format.
            using (FileStream docStream = File.OpenRead(MyDir + "Word document with missing file extension"))
            {
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(docStream);
                LoadFormat loadFormat = info.LoadFormat;

                Assert.AreEqual(LoadFormat.Doc, loadFormat);

                // Below are two methods of converting a LoadFormat to its corresponding SaveFormat.
                // 1 -  Get the file extension string for the LoadFormat, then get the corresponding SaveFormat from that string:
                string fileExtension = FileFormatUtil.LoadFormatToExtension(loadFormat);
                SaveFormat saveFormat = FileFormatUtil.ExtensionToSaveFormat(fileExtension);

                // 2 -  Convert the LoadFormat directly to its SaveFormat:
                saveFormat = FileFormatUtil.LoadFormatToSaveFormat(loadFormat);

                // Load a document from the stream, and then save it to the automatically detected file extension.
                Document doc = new Document(docStream);

                Assert.AreEqual(".doc", FileFormatUtil.SaveFormatToExtension(saveFormat));

                doc.Save(ArtifactsDir + "File.SaveToDetectedFileFormat" + FileFormatUtil.SaveFormatToExtension(saveFormat));
            }
            //ExEnd
        }

        [Test]
        public void DetectFileFormat_SaveFormatToLoadFormat()
        {
            //ExStart
            //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
            //ExSummary:Shows how to convert a save format to its corresponding load format.
            Assert.AreEqual(LoadFormat.Html, FileFormatUtil.SaveFormatToLoadFormat(SaveFormat.Html));

            // Some file types can have documents saved to, but not loaded from using Aspose.Words.
            // If we attempt to convert a save format of such a type to a load format, an exception will be thrown.
            Assert.Throws<ArgumentException>(() => FileFormatUtil.SaveFormatToLoadFormat(SaveFormat.Jpeg));
            //ExEnd
        }


        [Test]
        public void ExtractImages()
        {
            //ExStart
            //ExFor:Shape
            //ExFor:Shape.ImageData
            //ExFor:Shape.HasImage
            //ExFor:ImageData
            //ExFor:FileFormatUtil.ImageTypeToExtension(ImageType)
            //ExFor:ImageData.ImageType
            //ExFor:ImageData.Save(String)
            //ExFor:CompositeNode.GetChildNodes(NodeType, bool)
            //ExSummary:Shows how to extract images from a document, and save them to the local file system as individual files.
            Document doc = new Document(MyDir + "Images.docx");
            
            // Get the collection of shapes from the document,
            // and save the image data of every shape with an image as a file to the local file system.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            Assert.AreEqual(9, shapes.Count(s => ((Shape)s).HasImage));

            int imageIndex = 0;
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // The image data of shapes may contain images of many possible image formats. 
                    // We can determine a file extension for each image automatically, based on its format.
                    string imageFileName =
                        $"File.ExtractImages.{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                    shape.ImageData.Save(ArtifactsDir + imageFileName);
                    imageIndex++;
                }
            }
            //ExEnd

            Assert.AreEqual(9,Directory.GetFiles(ArtifactsDir).
                Count(s => Regex.IsMatch(s, @"^.+\.(jpeg|png|emf|wmf)$") && s.StartsWith(ArtifactsDir + "File.ExtractImages")));
        }
    }
}