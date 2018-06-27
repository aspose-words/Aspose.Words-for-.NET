using System;
using Aspose.Words;
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
            //ExSummary:Shows how to catch a FileCorrputedException
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
            // This will not work on a non-html file
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Document.doc");
            Assert.AreEqual(LoadFormat.Doc, info.LoadFormat);
            Assert.IsNull(info.Encoding);

            // This time the property will not be null
            info = FileFormatUtil.DetectFileFormat(MyDir + "Document.LoadFormat.html");
            Assert.AreEqual(LoadFormat.Html, info.LoadFormat);
            Assert.IsNotNull(info.Encoding);

            // It now has some more useful information
            Assert.AreEqual("iso-8859-1", info.Encoding.BodyName);
            Assert.AreEqual("Western European (Windows)", info.Encoding.EncodingName);
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

            Assert.AreEqual(LoadFormat.Docx, FileFormatUtil.ContentTypeToLoadFormat("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
            Assert.AreEqual(SaveFormat.Docx, FileFormatUtil.ContentTypeToSaveFormat("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));

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
    }
}
