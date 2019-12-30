// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
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
                Document doc = new Document(MyDir + "Corrupted.docx");
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
            // 'DetectFileFormat' not working on a non-html files
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Document.doc");
            Assert.AreEqual(LoadFormat.Doc, info.LoadFormat);
            Assert.IsNull(info.Encoding);

            // This time the property will not be null
            info = FileFormatUtil.DetectFileFormat(MyDir + "Document.LoadFormat.html");
            Assert.AreEqual(LoadFormat.Html, info.LoadFormat);
            Assert.IsNotNull(info.Encoding);

            // It now has some more useful information
            Assert.AreEqual("iso-8859-1", info.Encoding.BodyName);
            //ExEnd
        }

        [Test]
        public void FileFormatToString()
        {
            //ExStart
            //ExFor:FileFormatUtil.ContentTypeToLoadFormat(String)
            //ExFor:FileFormatUtil.ContentTypeToSaveFormat(String)
            //ExSummary:Shows how to find the corresponding Aspose load/save format from an IANA content type string.
            // Trying to search for a SaveFormat with a simple string will not work
            try
            {
                Assert.AreEqual(SaveFormat.Jpeg, FileFormatUtil.ContentTypeToSaveFormat("jpeg"));
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e.Message);
            }

            // The convertion methods only accept official IANA type names, which are all listed here:
            //      https://www.iana.org/assignments/media-types/media-types.xhtml
            // Note that if a corresponding SaveFormat or LoadFormat for a type from that list does not exist in the Aspose enums,
            // converting will raise an exception just like in the code above 

            // File types that can be saved to but not opened as documents will not have corresponding load formats
            // Attempting to convert them to load formats will raise an exception
            Assert.AreEqual(SaveFormat.Jpeg, FileFormatUtil.ContentTypeToSaveFormat("image/jpeg"));
            Assert.AreEqual(SaveFormat.Png, FileFormatUtil.ContentTypeToSaveFormat("image/png"));
            Assert.AreEqual(SaveFormat.Tiff, FileFormatUtil.ContentTypeToSaveFormat("image/tiff"));
            Assert.AreEqual(SaveFormat.Gif, FileFormatUtil.ContentTypeToSaveFormat("image/gif"));
            Assert.AreEqual(SaveFormat.Emf, FileFormatUtil.ContentTypeToSaveFormat("image/x-emf"));
            Assert.AreEqual(SaveFormat.Xps, FileFormatUtil.ContentTypeToSaveFormat("application/vnd.ms-xpsdocument"));
            Assert.AreEqual(SaveFormat.Pdf, FileFormatUtil.ContentTypeToSaveFormat("application/pdf"));
            Assert.AreEqual(SaveFormat.Svg, FileFormatUtil.ContentTypeToSaveFormat("image/svg+xml"));
            Assert.AreEqual(SaveFormat.Epub, FileFormatUtil.ContentTypeToSaveFormat("application/epub+zip"));

            // File types that can both be loaded and saved have corresponding load and save formats
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
        public void DetectFileFormat()
        {
            //ExStart
            //ExFor:FileFormatUtil.DetectFileFormat(String)
            //ExFor:FileFormatInfo
            //ExFor:FileFormatInfo.LoadFormat
            //ExFor:FileFormatInfo.IsEncrypted
            //ExFor:FileFormatInfo.HasDigitalSignature
            //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and other features of the document.
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Document.doc");
            Console.WriteLine("The document format is: " + FileFormatUtil.LoadFormatToExtension(info.LoadFormat));
            Console.WriteLine("Document is encrypted: " + info.IsEncrypted);
            Console.WriteLine("Document has a digital signature: " + info.HasDigitalSignature);
            //ExEnd
        }

        [Test]
        public void DetectFileFormat_EnumConversions()
        {
            //ExStart
            //ExFor:FileFormatUtil.DetectFileFormat(Stream)
            //ExFor:FileFormatUtil.LoadFormatToExtension(LoadFormat)
            //ExFor:FileFormatUtil.ExtensionToSaveFormat(String)
            //ExFor:FileFormatUtil.SaveFormatToExtension(SaveFormat)
            //ExFor:FileFormatUtil.LoadFormatToSaveFormat(LoadFormat)
            //ExFor:Document.OriginalFileName
            //ExFor:FileFormatInfo.LoadFormat
            //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document without any extension and save it with the correct file extension.
            // Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format
            // These are both times where you might need extract the file format as it's not visible
            // The file format of this document is actually ".doc"
            FileStream docStream = File.OpenRead(MyDir + "Document.FileWithoutExtension");
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(docStream);

            // Retrieve the LoadFormat of the document
            LoadFormat loadFormat = info.LoadFormat;

            // Let's show the different methods of converting LoadFormat enumerations to SaveFormat enumerations
            //
            // Method #1
            // Convert the LoadFormat to a String first for working with. The String will include the leading dot in front of the extension
            string fileExtension = FileFormatUtil.LoadFormatToExtension(loadFormat);
            // Now convert this extension into the corresponding SaveFormat enumeration
            SaveFormat saveFormat = FileFormatUtil.ExtensionToSaveFormat(fileExtension);

            // Method #2
            // Convert the LoadFormat enumeration directly to the SaveFormat enumeration
            saveFormat = FileFormatUtil.LoadFormatToSaveFormat(loadFormat);

            // Load a document from the stream.
            Document doc = new Document(docStream);

            // Save the document with the original file name, " Out" and the document's file extension
            doc.Save(ArtifactsDir + "Document.WithFileExtension" + FileFormatUtil.SaveFormatToExtension(saveFormat));
            //ExEnd

            Assert.AreEqual(".doc", FileFormatUtil.SaveFormatToExtension(saveFormat));
        }

        [Test]
        public void DetectFileFormat_SaveFormatToLoadFormat()
        {
            //ExStart
            //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
            //ExSummary:Shows how to use the FileFormatUtil class and to convert a SaveFormat enumeration into the corresponding LoadFormat enumeration.
            // Define the SaveFormat enumeration to convert
            const SaveFormat saveFormat = SaveFormat.Html;
            // Convert the SaveFormat enumeration to LoadFormat enumeration
            LoadFormat loadFormat = FileFormatUtil.SaveFormatToLoadFormat(saveFormat);
            Console.WriteLine("The converted LoadFormat is: " + FileFormatUtil.LoadFormatToExtension(loadFormat));
            //ExEnd

            Assert.AreEqual(".html", FileFormatUtil.SaveFormatToExtension(saveFormat));
            Assert.AreEqual(".html", FileFormatUtil.LoadFormatToExtension(loadFormat));
        }

        [Test]
        public void DetectDocumentSignatures()
        {
            //ExStart
            //ExFor:FileFormatUtil.DetectFileFormat(String)
            //ExFor:FileFormatInfo.HasDigitalSignature
            //ExSummary:Shows how to check a document for digital signatures before loading it into a Document object.
            // The path to the document which is to be processed
            string filePath = MyDir + "Document.Signed.docx";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            if (info.HasDigitalSignature)
            {
                Console.WriteLine(
                    "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                    Path.GetFileName(filePath));
            }
            //ExEnd
        }

        //ExStart
        //ExFor:Shape
        //ExFor:Shape.ImageData
        //ExFor:Shape.HasImage
        //ExFor:ImageData
        //ExFor:FileFormatUtil.ImageTypeToExtension(ImageType)
        //ExFor:ImageData.ImageType
        //ExFor:ImageData.Save(String)
        //ExFor:CompositeNode.GetChildNodes(NodeType, bool)
        //ExSummary:Shows how to extract images from a document and save them as files.
        [Test] //ExSkip
        public void ExtractImagesToFiles()
        {
            Document doc = new Document(MyDir + "Image.SampleImages.doc");

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string imageFileName =
                        $"Image.ExportImages.{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                    shape.ImageData.Save(ArtifactsDir + imageFileName);
                    imageIndex++;
                }
            }
        }
        //ExEnd
    }
}