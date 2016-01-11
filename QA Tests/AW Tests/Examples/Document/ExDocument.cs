// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Web;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using Aspose.Words.Tables;
using NUnit.Framework;
using QA_Tests.Tests;
#if !JAVA
//ExStart
//ExId:ImportForDigitalSignatures
//ExSummary:The import required to use the X509Certificate2 class.
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
//ExEnd
#endif

namespace QA_Tests.Examples.Document
{
    [TestFixture]
    public class ExDocument : QaTestsBase
    {
        [Test]
        public void LicenseFromFileNoPath()
        {
            // Copy a license to the bin folder so the example can execute.
            string dstFileName = Path.Combine(AssemblyDir, "Aspose.Words.lic");
            File.Copy(TestLicenseFileName, dstFileName);

            //ExStart
            //ExFor:License
            //ExFor:License.#ctor
            //ExFor:License.SetLicense(String)
            //ExId:LicenseFromFileNoPath
            //ExSummary:In this example Aspose.Words will attempt to find the license file in the embedded resources or in the assembly folders.
            Aspose.Words.License license = new Aspose.Words.License();
            license.SetLicense("Aspose.Words.lic");
            //ExEnd

            // Cleanup by removing the license.
            license.SetLicense("");
            File.Delete(dstFileName);
        }

        [Test]
        public void LicenseFromStream()
        {
            Stream myStream = File.OpenRead(TestLicenseFileName);
            try
            {
                //ExStart
                //ExFor:License.SetLicense(Stream)
                //ExId:LicenseFromStream
                //ExSummary:Initializes a license from a stream.
                Aspose.Words.License license = new Aspose.Words.License();
                license.SetLicense(myStream);
                //ExEnd
            }
            finally
            {
                myStream.Close();
            }
        }

        [Test]
        public void DocumentCtor()
        {
            //ExStart
            //ExId:DocumentCtor
            //ExSummary:Shows how to create a blank document. Note the blank document contains one section and one paragraph.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            //ExEnd
        }

        [Test]
        public void OpenFromFile()
        {
            //ExStart
            //ExFor:Document.#ctor(String)
            //ExId:OpenFromFile
            //ExSummary:Opens a document from a file.
            // Open a document. The file is opened read only and only for the duration of the constructor.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            //ExEnd

            //ExStart
            //ExFor:Document.Save(String)
            //ExId:SaveToFile
            //ExSummary:Saves a document to a file.
            doc.Save(ExDir + "Document.OpenFromFile Out.doc");
            //ExEnd
        }

        [Test]
        public void OpenAndSaveToFile()
        {
            //ExStart
            //ExId:OpenAndSaveToFile
            //ExSummary:Opens a document from a file and saves it to a different format
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            doc.Save(ExDir + "Document Out.html");
            //ExEnd
        }

        [Test]
        public void OpenFromStream()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream)
            //ExId:OpenFromStream
            //ExSummary:Opens a document from a stream.
            // Open the stream. Read only access is enough for Aspose.Words to load a document.
            Stream stream = File.OpenRead(ExDir + "Document.doc");

            // Load the entire document into memory.
            Aspose.Words.Document doc = new Aspose.Words.Document(stream);

            // You can close the stream now, it is no longer needed because the document is in memory.
            stream.Close();

            // ... do something with the document
            //ExEnd

            Assert.AreEqual("Hello World!\x000c", doc.GetText());
        }

        [Test]
        public void OpenFromStreamWithBaseUri()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream,LoadOptions)
            //ExFor:LoadOptions
            //ExFor:LoadOptions.BaseUri
            //ExId:DocumentCtor_LoadOptions
            //ExSummary:Opens an HTML document with images from a stream using a base URI.

            // We are opening this HTML file:      
            //    <html>
            //    <body>
            //    <p>Simple file.</p>
            //    <p><img src="Aspose.Words.gif" width="80" height="60"></p>
            //    </body>
            //    </html>
            string fileName = ExDir + "Document.OpenFromStreamWithBaseUri.html";

            // Open the stream.
            Stream stream = File.OpenRead(fileName);

            // Open the document. Note the Document constructor detects HTML format automatically.
            // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.BaseUri = ExDir;
            Aspose.Words.Document doc = new Aspose.Words.Document(stream, loadOptions);

            // You can close the stream now, it is no longer needed because the document is in memory.
            stream.Close();

            // Save in the DOC format.
            doc.Save(ExDir + "Document.OpenFromStreamWithBaseUri Out.doc");
            //ExEnd

            // Lets make sure the image was imported successfully into a Shape node.
            // Get the first shape node in the document.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            // Verify some properties of the image.
            Assert.IsTrue(shape.IsImage);
            Assert.IsNotNull(shape.ImageData.ImageBytes);
            Assert.AreEqual(80.0, Aspose.Words.ConvertUtil.PointToPixel(shape.Width));
            Assert.AreEqual(60.0, Aspose.Words.ConvertUtil.PointToPixel(shape.Height));
        }

        [Test]
        public void OpenDocumentFromWeb()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream)
            //ExSummary://ExSummary:Retrieves a document from a URL and saves it to disk in a different format.
            // This is the URL address pointing to where to find the document.
            string url = "http://www.aspose.com/demos/.net-components/aspose.words/csharp/general/Common/Documents/DinnerInvitationDemo.doc";

            // The easiest way to load our document from the internet is make use of the 
            // System.Net.WebClient class. Create an instance of it and pass the URL
            // to download from.
            WebClient webClient = new WebClient();

            // Download the bytes from the location referenced by the URL.
            byte[] dataBytes = webClient.DownloadData(url);

            // Wrap the bytes representing the document in memory into a MemoryStream object.
            MemoryStream byteStream = new MemoryStream(dataBytes);

            // Load this memory stream into a new Aspose.Words Document.
            // The file format of the passed data is inferred from the content of the bytes itself. 
            // You can load any document format supported by Aspose.Words in the same way.
            Aspose.Words.Document doc = new Aspose.Words.Document(byteStream);

            // Convert the document to any format supported by Aspose.Words.
            doc.Save(ExDir + "Document.OpenFromWeb Out.docx");
            //ExEnd
        }

        [Test]
        public void InsertHtmlFromWebPage()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream, LoadOptions)
            //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
            //ExFor:LoadFormat
            //ExSummary:Shows how to insert the HTML conntents from a web page into a new document.
            // The url of the page to load 
            String url = "http://www.aspose.com/";

            // Create a WebClient object to easily extract the HTML from the page.
            WebClient client = new WebClient();
            String pageSource = client.DownloadString(url);
            client.Dispose();

            // Get the HTML as bytes for loading into a stream.
            Encoding encoding = client.Encoding;
            byte[] pageBytes = encoding.GetBytes(pageSource);

            // Load the HTML into a stream.
            MemoryStream stream = new MemoryStream(pageBytes);

            // The baseUri property should be set to ensure any relative img paths are retrieved correctly.
            LoadOptions options = new LoadOptions(Aspose.Words.LoadFormat.Html, "", url);

            // Load the HTML document from stream and pass the LoadOptions object.
            Aspose.Words.Document doc = new Aspose.Words.Document(stream, options);

            // Save the document to disk.
            // The extension of the filename can be changed to save the document into other formats. e.g PDF, DOCX, ODT, RTF.
            doc.Save(ExDir + "Document.HtmlPageFromWebpage Out.doc");
            //ExEnd
        }

        [Test]
        public void LoadFormat()
        {
            //ExStart
            //ExFor:Document.#ctor(String,LoadOptions)
            //ExFor:LoadFormat
            //ExSummary:Explicitly loads a document as HTML without automatic file format detection.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFormat = Aspose.Words.LoadFormat.Html;
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.LoadFormat.html", loadOptions);
            //ExEnd
        }

        [Test]
        public void LoadEncryptedFromFile()
        {
            //ExStart
            //ExFor:Document.#ctor(String,LoadOptions)
            //ExFor:LoadOptions
            //ExFor:LoadOptions.#ctor(String)
            //ExId:OpenEncrypted
            //ExSummary:Loads a Microsoft Word document encrypted with a password.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.LoadEncrypted.doc", new LoadOptions("qwerty"));
            //ExEnd
        }

        [Test]
        public void LoadEncryptedFromStream()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream,LoadOptions)
            //ExSummary:Loads a Microsoft Word document encrypted with a password from a stream.
            Stream stream = File.OpenRead(ExDir + "Document.LoadEncrypted.doc");
            Aspose.Words.Document doc = new Aspose.Words.Document(stream, new LoadOptions("qwerty"));
            stream.Close();
            //ExEnd
        }

        [Test]
        public void ConvertToHtml()
        {
            //ExStart
            //ExFor:Document.Save(String,SaveFormat)
            //ExFor:SaveFormat
            //ExSummary:Converts from DOC to HTML format.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Save(ExDir + "Document.ConvertToHtml Out.html", SaveFormat.Html);
            //ExEnd
        }

        [Test]
        public void ConvertToMhtml()
        {
            //ExStart
            //ExFor:Document.Save(String)
            //ExSummary:Converts from DOC to MHTML format.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Save(ExDir + "Document.ConvertToMhtml Out.mht");
            //ExEnd
        }

        [Test]
        public void ConvertToTxt()
        {
            //ExStart
            //ExId:ExtractContentSaveAsText
            //ExSummary:Shows how to save a document in TXT format.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Save(ExDir + "Document.ConvertToTxt Out.txt");
            //ExEnd
        }

        [Test]
        public void Doc2PdfSave()
        {
            //ExStart
            //ExFor:Document
            //ExFor:Document.Save(String)
            //ExId:Doc2PdfSave
            //ExSummary:Converts a whole document from DOC to PDF using default options.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Save(ExDir + "Document.Doc2PdfSave Out.pdf");
            //ExEnd
        }

        [Test]
        public void SaveToStream()
        {
            //ExStart
            //ExFor:Document.Save(Stream,SaveFormat)
            //ExId:SaveToStream
            //ExSummary:Shows how to save a document to a stream.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            // Rewind the stream position back to zero so it is ready for next reader.
            dstStream.Position = 0;
            //ExEnd
        }

        /// <summary>
        /// RK We are not actually executing this as a test because it does not seem to work without ASP.NET
        /// </summary>
        public void SaveToBrowser()
        {
            // Create a dummy HTTP response.
            HttpResponse Response = new HttpResponse(null);

            //ExStart
            //ExId:SaveToBrowser
            //ExSummary:Shows how to send a document to the client browser from an ASP.NET code.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            doc.Save(Response, "Report Out.doc", ContentDisposition.Inline, null);
            //ExEnd
        }

        [Test]
        public void Doc2EpubSave()
        {
            //ExStart
            //ExId:Doc2EpubSave
            //ExSummary:Converts a document to EPUB using default save options.

            // Open an existing document from disk.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.EpubConversion.doc");

            // Save the document in EPUB format.
            doc.Save(ExDir + "Document.EpubConversion Out.epub");
            //ExEnd
        }

        [Test]
        public void Doc2EpubSaveWithOptions()
        {
            //ExStart
            //ExFor:HtmlSaveOptions
            //ExFor:HtmlSaveOptions.#ctor
            //ExFor:HtmlSaveOptions.Encoding
            //ExFor:HtmlSaveOptions.DocumentSplitCriteria
            //ExFor:HtmlSaveOptions.ExportDocumentProperties
            //ExFor:HtmlSaveOptions.SaveFormat
            //ExId:Doc2EpubSaveWithOptions
            //ExSummary:Converts a document to EPUB with save options specified.
            // Open an existing document from disk.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.EpubConversion.doc");

            // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
            // how the output document is saved.
            HtmlSaveOptions saveOptions =
                new HtmlSaveOptions();

            // Specify the desired encoding.
            saveOptions.Encoding = System.Text.Encoding.UTF8;

            // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
            // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
            // HTML files greater than a certain size e.g 300kb.
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

            // Specify that we want to export document properties.
            saveOptions.ExportDocumentProperties = true;

            // Specify that we want to save in EPUB format.
            saveOptions.SaveFormat = SaveFormat.Epub;

            // Export the document as an EPUB file.
            doc.Save(ExDir + "Document.EpubConversion Out.epub", saveOptions);
            //ExEnd
        }

        [Test]
        public void SaveHtmlPrettyFormat()
        {
            //ExStart
            //ExFor:SaveOptions.PrettyFormat
            //ExSummary:Shows how to pass an option to export HTML tags in a well spaced, human readable format.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Enabling the PrettyFormat setting will export HTML in an indented format that is easy to read.
            // If this is setting is false (by default) then the HTML tags will be exported in condensed form with no indentation.
            htmlOptions.PrettyFormat = true;

            doc.Save(ExDir + "Document.PrettyFormat Out.html", htmlOptions);
            //ExEnd
        }

        [Test]
        public void SaveHtmlWithOptions()
        {
            //ExStart
            //ExFor:HtmlSaveOptions
            //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
            //ExFor:HtmlSaveOptions.ImagesFolder
            //ExId:SaveWithOptions
            //ExSummary:Shows how to set save options before saving a document to HTML.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Rendering.doc");
            
            // This is the directory we want the exported images to be saved to.
            string imagesDir = Path.Combine(ExDir, "Images");
            
            // The folder specified needs to exist and should be empty.
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportTextInputFormFieldAsText = true;
            options.ImagesFolder = imagesDir;

            doc.Save(ExDir + "Document.SaveWithOptions Out.html", options);
            //ExEnd

            // Verify the images were saved to the correct location.
            Assert.IsTrue(File.Exists(ExDir + "Document.SaveWithOptions Out.html"));
            Assert.AreEqual(9, Directory.GetFiles(imagesDir).Length);
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void SaveHtmlExportFontsCaller()
        {
            SaveHtmlExportFonts();
        }

        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontResources
        //ExFor:HtmlSaveOptions.FontSavingCallback
        //ExFor:IFontSavingCallback
        //ExFor:IFontSavingCallback.FontSaving
        //ExFor:FontSavingArgs
        //ExFor:FontSavingArgs.FontFamilyName
        //ExFor:FontSavingArgs.FontFileName
        //ExId:SaveHtmlExportFonts
        //ExSummary:Shows how to define custom logic for handling font exporting when saving to HTML based formats.
        public void SaveHtmlExportFonts()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            // Set the option to export font resources.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml);
            options.ExportFontResources = true;
            // Create and pass the object which implements the handler methods.
            options.FontSavingCallback = new HandleFontSaving();

            doc.Save(ExDir + "Document.SaveWithFontsExport Out.html", options);
        }

        public class HandleFontSaving : IFontSavingCallback
        {
            void IFontSavingCallback.FontSaving(FontSavingArgs args)
            {
                // You can implement logic here to rename fonts, save to file etc. For this example just print some details about the current font being handled.
                Console.WriteLine("Font Name = {0}, Font Filename = {1}", args.FontFamilyName, args.FontFileName);
            }
        }
        //ExEnd

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void SaveHtmlExportImagesCaller()
        {
            SaveHtmlExportImages();
        }

        //ExStart
        //ExFor:IImageSavingCallback
        //ExFor:IImageSavingCallback.ImageSaving
        //ExFor:ImageSavingArgs
        //ExFor:ImageSavingArgs.ImageFileName
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.ImageSavingCallback
        //ExId:SaveHtmlCustomExport
        //ExSummary:Shows how to define custom logic for controlling how images are saved when exporting to HTML based formats.
        public void SaveHtmlExportImages()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            // Create and pass the object which implements the handler methods.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ImageSavingCallback = new HandleImageSaving();

            doc.Save(ExDir + "Document.SaveWithCustomImagesExport Out.html", options);
        }

        public class HandleImageSaving : IImageSavingCallback
        {
            void IImageSavingCallback.ImageSaving(ImageSavingArgs e)
            {
                // Change any images in the document being exported with the extension of "jpeg" to "jpg".
                if (e.ImageFileName.EndsWith(".jpeg"))
                    e.ImageFileName = e.ImageFileName.Replace(".jpeg", ".jpg");
            }
        }
        //ExEnd

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void TestNodeChangingInDocumentCaller()
        {
            TestNodeChangingInDocument();
        }

        //ExStart
        //ExFor:INodeChangingCallback
        //ExFor:INodeChangingCallback.NodeInserting
        //ExFor:INodeChangingCallback.NodeInserted
        //ExFor:INodeChangingCallback.NodeRemoving
        //ExFor:INodeChangingCallback.NodeRemoved
        //ExFor:NodeChangingArgs
        //ExFor:NodeChangingArgs.Node
        //ExFor:DocumentBase.NodeChangingCallback
        //ExId:NodeChangingInDocument
        //ExSummary:Shows how to implement custom logic over node insertion in the document by changing the font of inserted HTML content.
        public void TestNodeChangingInDocument()
        {
            // Create a blank document object
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set up and pass the object which implements the handler methods.
            doc.NodeChangingCallback = new HandleNodeChanging_FontChanger();

            // Insert sample HTML content
            builder.InsertHtml("<p>Hello World</p>");

            doc.Save(ExDir + "Document.FontChanger Out.doc");

            // Check that the inserted content has the correct formatting
            Run run = (Run)doc.GetChild(NodeType.Run, 0, true);
            Assert.AreEqual(24.0, run.Font.Size);
            Assert.AreEqual("Arial", run.Font.Name);
        }

        public class HandleNodeChanging_FontChanger : INodeChangingCallback
        {
            // Implement the NodeInserted handler to set default font settings for every Run node inserted into the Document
            void INodeChangingCallback.NodeInserted(NodeChangingArgs args)
            {
                // Change the font of inserted text contained in the Run nodes.
                if (args.Node.NodeType == NodeType.Run)
                {
                    Aspose.Words.Font font = ((Run)args.Node).Font;
                    font.Size = 24;
                    font.Name = "Arial";
                }
            }

            void INodeChangingCallback.NodeInserting(NodeChangingArgs args)
            {
                // Do Nothing
            }

            void INodeChangingCallback.NodeRemoved(NodeChangingArgs args)
            {
                // Do Nothing
            }

            void INodeChangingCallback.NodeRemoving(NodeChangingArgs args)
            {
                // Do Nothing
            }
        }
        //ExEnd

        [Test]
        public void DetectFileFormat()
        {
            //ExStart
            //ExFor:FileFormatUtil.DetectFileFormat(String)
            //ExFor:FileFormatInfo
            //ExFor:FileFormatInfo.LoadFormat
            //ExFor:FileFormatInfo.IsEncrypted
            //ExFor:FileFormatInfo.HasDigitalSignature
            //ExId:DetectFileFormat
            //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and other features of the document.
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(ExDir + "Document.doc");
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
            // Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format. These are both times where you might need extract the file format as it's not visible
            FileStream docStream = File.OpenRead(ExDir + "Document.FileWithoutExtension"); // The file format of this document is actually ".doc"
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(docStream);

            // Retrieve the LoadFormat of the document.
            LoadFormat loadFormat = info.LoadFormat;

            // Let's show the different methods of converting LoadFormat enumerations to SaveFormat enumerations.
            //
            // Method #1
            // Convert the LoadFormat to a string first for working with. The string will include the leading dot in front of the extension.
            string fileExtension = FileFormatUtil.LoadFormatToExtension(loadFormat);
            // Now convert this extension into the corresponding SaveFormat enumeration
            SaveFormat saveFormat = FileFormatUtil.ExtensionToSaveFormat(fileExtension);

            // Method #2
            // Convert the LoadFormat enumeration directly to the SaveFormat enumeration.
            saveFormat = FileFormatUtil.LoadFormatToSaveFormat(loadFormat);

            // Load a document from the stream.
            Aspose.Words.Document doc = new Aspose.Words.Document(docStream);

            // Save the document with the original file name, " Out" and the document's file extension.
            doc.Save(ExDir + "Document.WithFileExtension Out" + FileFormatUtil.SaveFormatToExtension(saveFormat));
            //ExEnd

            Assert.AreEqual(".doc", FileFormatUtil.SaveFormatToExtension(saveFormat));
        }

        [Test]
        public void DetectFileFormat_SaveFormatToLoadFormat()
        {
            //ExStart
            //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
            //ExSummary:Shows how to use the FileFormatUtil class and to convert a SaveFormat enumeration into the corresponding LoadFormat enumeration.
            // Define the SaveFormat enumeration to convert.
            SaveFormat saveFormat = SaveFormat.Html;
            // Convert the SaveFormat enumeration to LoadFormat enumeration.
            LoadFormat loadFormat = FileFormatUtil.SaveFormatToLoadFormat(saveFormat);
            Console.WriteLine("The converted LoadFormat is: " + FileFormatUtil.LoadFormatToExtension(loadFormat));
            //ExEnd
            
            Assert.AreEqual(".html", FileFormatUtil.SaveFormatToExtension(saveFormat));
            Assert.AreEqual(".html", FileFormatUtil.LoadFormatToExtension(loadFormat)) ;
        }

        [Test]
        public void AppendDocument()
        {
            //ExStart
            //ExFor:Document.AppendDocument(Document, ImportFormatMode)
            //ExSummary:Shows how to append a document to the end of another document.
            // The document that the content will be appended to.
            Aspose.Words.Document dstDoc = new Aspose.Words.Document(ExDir + "Document.doc");
            // The document to append.
            Aspose.Words.Document srcDoc = new Aspose.Words.Document(ExDir + "DocumentBuilder.doc");

            // Append the source document to the destination document.
            // Pass format mode to retain the original formatting of the source document when importing it.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the document.
            dstDoc.Save(ExDir + "Document.AppendDocument Out.doc");
            //ExEnd
        }

        [Test]
        // Using this file path keeps the example making sense when compared with automation so we expect
        // the file not to be found.
        [ExpectedException(typeof(System.IO.FileNotFoundException))]
        public void AppendDocumentFromAutomation()
        {
            //ExStart
            //ExId:AppendDocumentFromAutomation
            //ExSummary:Shows how to join multiple documents together.
            // The document that the other documents will be appended to.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            // We should call this method to clear this document of any existing content.
            doc.RemoveAllChildren();

            int recordCount = 5;
            for (int i = 1; i <= recordCount; i++)
            {
                // Open the document to join.
                Aspose.Words.Document srcDoc = new Aspose.Words.Document(@"C:\DetailsList.doc");

                // Append the source document at the end of the destination document.
                doc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

                // In automation you were required to insert a new section break at this point, however in Aspose.Words we 
                // don't need to do anything here as the appended document is imported as separate sectons already.

                // If this is the second document or above being appended then unlink all headers footers in this section 
                // from the headers and footers of the previous section.
                if (i > 1)
                    doc.Sections[i].HeadersFooters.LinkToPrevious(false);
            }
            //ExEnd
        }

        [Test]
        public void DetectDocumentSignatures()
        {
            //ExStart
            //ExFor:FileFormatUtil.DetectFileFormat(String)
            //ExFor:FileFormatInfo.HasDigitalSignature
            //ExId:DetectDocumentSignatures
            //ExSummary:Shows how to check a document for digital signatures before loading it into a Document object.
            // The path to the document which is to be processed.
            string filePath = ExDir + "Document.Signed.docx";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            if (info.HasDigitalSignature)
            {
                Console.WriteLine(string.Format("Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.", Path.GetFileName(filePath)));
            }
            //ExEnd
        }

        [Test]
        public void ValidateAllDocumentSignatures()
        {
            //ExStart
            //ExFor:Document.DigitalSignatures
            //ExFor:DigitalSignatureCollection
            //ExFor:DigitalSignatureCollection.IsValid
            //ExId:ValidateAllDocumentSignatures
            //ExSummary:Shows how to validate all signatures in a document.
            // Load the signed document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.Signed.docx");

            if (doc.DigitalSignatures.IsValid)
                Console.WriteLine("Signatures belonging to this document are valid");
            else
                Console.WriteLine("Signatures belonging to this document are NOT valid");
            //ExEnd

            Assert.True(doc.DigitalSignatures.IsValid);
        }

        [Test]
        public void ValidateIndividualDocumentSignatures()
        {
            //ExStart
            //ExFor:DigitalSignature
            //ExFor:Document.DigitalSignatures
            //ExFor:DigitalSignature.IsValid
            //ExFor:DigitalSignature.Comments
            //ExFor:DigitalSignature.SignTime
            //ExFor:DigitalSignature.SignatureType
            //ExFor:DigitalSignature.Certificate
            //ExId:ValidateIndividualSignatures
            //ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
            // Load the document which contains signature.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.Signed.docx");

            foreach (Aspose.Words.DigitalSignature signature in doc.DigitalSignatures)
            {
                Console.WriteLine("*** Signature Found ***");
                Console.WriteLine("Is valid: " + signature.IsValid);
                Console.WriteLine("Reason for signing: " + signature.Comments); // This property is available in MS Word documents only.
                Console.WriteLine("Signature type: " + signature.SignatureType.ToString());
                Console.WriteLine("Time of signing: " + signature.SignTime);
                Console.WriteLine("Subject name: " + signature.Certificate.SubjectName.ToString());
                Console.WriteLine("Issuer name: " + signature.Certificate.IssuerName.Name);
                Console.WriteLine();
            }
            //ExEnd

            Aspose.Words.DigitalSignature digitalSig = doc.DigitalSignatures[0];
            Assert.True(digitalSig.IsValid);
            Assert.AreEqual("Test Sign", digitalSig.Comments);
            Assert.AreEqual("XmlDsig", digitalSig.SignatureType.ToString());
            Assert.True(digitalSig.Certificate.Subject.Contains("Aspose Pty Ltd"));
            Assert.True(digitalSig.Certificate.IssuerName.Name.Contains("VeriSign"));
        }

        [Test]
        // We don't include a sample certificate with the examples
        // so this exception is expected instead since the file is not there.
        [ExpectedException(typeof(System.Security.Cryptography.CryptographicException))]
        public void SignPDFDocument()
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:PdfDigitalSignatureDetails
            //ExFor:PdfSaveOptions.DigitalSignatureDetails
            //ExFor:PdfDigitalSignatureDetails.#ctor(X509Certificate2, String, String, DateTime)
            //ExId:SignPDFDocument
            //ExSummary:Shows how to sign a generated PDF document using Aspose.Words.
            // Create a simple document from scratch.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Test Signed PDF.");

            // Load the certificate from disk.
            // The other constructor overloads can be used to load certificates from different locations.
            X509Certificate2 cert = new X509Certificate2(
                ExDir + "certificate.pfx", "feyb4lgcfbme");

            // Pass the certificate and details to the save options class to sign with.
            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(
                cert,
                "Test Signing",
                "Aspose Office",
                DateTime.Now);

            // Save the document as PDF with the digital signature set.
            doc.Save(ExDir + "Document.Signed Out.pdf", options);
            //ExEnd
        }

        [Test]
        public void AppendAllDocumentsInFolder()
        {
            // Delete the file that was created by the previous run as I don't want to append it again.
            File.Delete(ExDir + "Document.AppendDocumentsFromFolder Out.doc");

            //ExStart
            //ExFor:Document.AppendDocument(Document, ImportFormatMode)
            //ExSummary:Shows how to use the AppendDocument method to combine all the documents in a folder to the end of a template document.
            // Lets start with a simple template and append all the documents in a folder to this document.
            Aspose.Words.Document baseDoc = new Aspose.Words.Document();

            // Add some content to the template.
            DocumentBuilder builder = new DocumentBuilder(baseDoc);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Template Document");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Some content here");

            // Gather the files which will be appended to our template document.
            // In this case we add the optional parameter to include the search only for files with the ".doc" extension.
            ArrayList files = new ArrayList(Directory.GetFiles(ExDir, "*.doc"));
            // The list of files may come in any order, let's sort the files by name so the documents are enumerated alphabetically.
            files.Sort();

            // Iterate through every file in the directory and append each one to the end of the template document.
            foreach (string fileName in files)
            {
                // We have some encrypted test documents in our directory, Aspose.Words can open encrypted documents 
                // but only with the correct password. Let's just skip them here for simplicity.
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(fileName);
                if (info.IsEncrypted)
                    continue;

                Aspose.Words.Document subDoc = new Aspose.Words.Document(fileName);
                baseDoc.AppendDocument(subDoc, ImportFormatMode.UseDestinationStyles);
            }

            // Save the combined document to disk.
            baseDoc.Save(ExDir + "Document.AppendDocumentsFromFolder Out.doc");
            //ExEnd
        }

        [Test]
        public void JoinRunsWithSameFormatting()
        {
            //ExStart
            //ExFor:Document.JoinRunsWithSameFormatting
            //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
            // Let's load this particular document. It contains a lot of content that has been edited many times.
            // This means the document will most likely contain a large number of runs with duplicate formatting.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Rendering.doc");

            // This is for illustration purposes only, remember how many run nodes we had in the original document.
            int runsBefore = doc.GetChildNodes(NodeType.Run, true).Count;

            // Join runs with the same formatting. This is useful to speed up processing and may also reduce redundant
            // tags when exporting to HTML which will reduce the output file size.
            int joinCount = doc.JoinRunsWithSameFormatting();

            // This is for illustration purposes only, see how many runs are left after joining.
            int runsAfter = doc.GetChildNodes(NodeType.Run, true).Count;

            Console.WriteLine("Number of runs before:{0}, after:{1}, joined:{2}", runsBefore, runsAfter, joinCount);

            // Save the optimized document to disk.
            doc.Save(ExDir + "Document.JoinRunsWithSameFormatting Out.html");
            //ExEnd

            // Verify that runs were joined in the document.
            Assert.Less(runsAfter, runsBefore);
            Assert.AreNotEqual(0, joinCount);
        }

        [Test]
        public void DetachTemplate()
        {
            //ExStart
            //ExFor:Document.AttachedTemplate
            //ExSummary:Opens a document, makes sure it is no longer attached to a template and saves the document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            doc.AttachedTemplate = "";
            doc.Save(ExDir + "Document.DetachTemplate Out.doc");
            //ExEnd
        }

        [Test]
        public void DefaultTabStop()
        {
            //ExStart
            //ExFor:Document.DefaultTabStop
            //ExFor:ControlChar.Tab
            //ExFor:ControlChar.TabChar
            //ExSummary:Changes default tab positions for the document and inserts text with some tab characters.
            DocumentBuilder builder = new DocumentBuilder();

            // Set default tab stop to 72 points (1 inch).
            builder.Document.DefaultTabStop = 72;

            builder.Writeln("Hello" + ControlChar.Tab + "World!");
            builder.Writeln("Hello" + ControlChar.TabChar + "World!");
            //ExEnd
        }

        [Test]
        public void CloneDocument()
        {
            //ExStart
            //ExFor:Document.Clone
            //ExId:CloneDocument
            //ExSummary:Shows how to deep clone a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.Document clone = doc.Clone();
            //ExEnd
        }

        [Test]
        public void ChangeFieldUpdateCultureSource()
        {
            // We will test this functionality creating a document with two fields with date formatting
            // field where the set language is different than the current culture, e.g German.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert content with German locale.
            builder.Font.LocaleId = 1031;
            builder.InsertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
            builder.Write(" - ");
            builder.InsertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

            // Make sure that English culture is set then execute mail merge using current culture for
            // date formatting.
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            doc.MailMerge.Execute(new string[] { "Date1" }, new object[] { new DateTime(2011, 1, 01) });

            //ExStart
            //ExFor:Document.FieldOptions
            //ExFor:FieldOptions
            //ExFor:FieldOptions.FieldUpdateCultureSource
            //ExFor:FieldUpdateCultureSource
            //ExId:ChangeFieldUpdateCultureSource
            //ExSummary:Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
            // Set the culture used during field update to the culture used by the field.
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.MailMerge.Execute(new string[] { "Date2" }, new object[] { new DateTime(2011, 1, 01) });
            //ExEnd

            // Verify the field update behaviour is correct.
            Assert.AreEqual("Saturday, 1 January 2011 - Samstag, 1 Januar 2011", doc.Range.Text.Trim());

            // Restore the original culture.
            Thread.CurrentThread.CurrentCulture = currentCulture;
        }

        [Test]
        public void ControlListLabelsExportToHtml()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Lists.PrintOutAllLists.doc");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // This option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
            // otherwise HTML <p> tag is used. This is also the default value.
            saveOptions.ExportListLabels = ExportListLabels.Auto;
            doc.Save(ExDir + "Document.ExportListLabels Auto Out.html", saveOptions);

            // Using this option the <p> tag is used for any list label representation.
            saveOptions.ExportListLabels = ExportListLabels.AsInlineText;
            doc.Save(ExDir + "Document.ExportListLabels InlineText Out.html", saveOptions);

            // The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
            saveOptions.ExportListLabels = ExportListLabels.ByHtmlTags;
            doc.Save(ExDir + "Document.ExportListLabels HtmlTags Out.html", saveOptions);
        }

        [Test]
        public void DocumentGetText_ToString()
        {
            //ExStart
            //ExFor:CompositeNode.GetText
            //ExFor:Node.ToString(SaveFormat)
            //ExId:NodeTxtExportDifferences
            //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
            Aspose.Words.Document doc = new Aspose.Words.Document();

            // Enter a dummy field into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Field");

            // GetText will retrieve all field codes and special characters
            Console.WriteLine("GetText() Result: " + doc.GetText());

            // ToString will export the node to the specified format. When converted to text it will not retrieve fields code 
            // or special characters, but will still contain some natural formatting characters such as paragraph markers etc. 
            // This is the same as "viewing" the document as if it was opened in a text editor.
            Console.WriteLine("ToString() Result: " + doc.ToString(SaveFormat.Text));
            //ExEnd
        }

        [Test]
        public void DocumentByteArray()
        {
            //ExStart
            //ExId:DocumentToFromByteArray
            //ExSummary:Shows how to convert a document object to an array of bytes and back into a document object again.
            // Load the document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            // Create a new memory stream.
            MemoryStream outStream = new MemoryStream();
            // Save the document to stream.
            doc.Save(outStream, SaveFormat.Docx);

            // Convert the document to byte form.
            byte[] docBytes = outStream.ToArray();

            // The bytes are now ready to be stored/transmitted.

            // Now reverse the steps to load the bytes back into a document object.
            MemoryStream inStream = new MemoryStream(docBytes);

            // Load the stream into a new document object.
            Aspose.Words.Document loadDoc = new Aspose.Words.Document(inStream);
            //ExEnd

            Assert.AreEqual(doc.GetText(), loadDoc.GetText());
        }

        [Test]
        public void ProtectUnprotectDocument()
        {
            //ExStart
            //ExFor:Document.Protect(ProtectionType,String)
            //ExId:ProtectDocument
            //ExSummary:Shows how to protect a document.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");
            //ExEnd

            //ExStart
            //ExFor:Document.Unprotect
            //ExId:UnprotectDocument
            //ExSummary:Shows how to unprotect any document. Note that the password is not required.
            doc.Unprotect();
            //ExEnd
        }

        [Test]
        public void GetProtectionType()
        {
            //ExStart
            //ExFor:Document.ProtectionType
            //ExId:GetProtectionType
            //ExSummary:Shows how to get protection type currently set in the document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            ProtectionType protectionType = doc.ProtectionType;
            //ExEnd
        }

        [Test]
        public void DocumentEnsureMinimum()
        {
            //ExStart
            //ExFor:Document.EnsureMinimum
            //ExSummary:Shows how to ensure the Document is valid (has the minimum nodes required to be valid).
            // Create a blank document then remove all nodes from it, the result will be a completely empty document.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            doc.RemoveAllChildren();

            // Ensure that the document is valid. Since the document has no nodes this method will create an empty section
            // and add an empty paragraph to make it valid.
            doc.EnsureMinimum();
            //ExEnd
        }

        [Test]
        public void RemoveMacrosFromDocument()
        {
            //ExStart
            //ExFor:Document.RemoveMacros
            //ExSummary:Shows how to remove all macros from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            doc.RemoveMacros();
            //ExEnd
        }

        [Test]
        public void UpdateTableLayout()
        {
            //ExStart
            //ExFor:Document.UpdateTableLayout
            //ExId:UpdateTableLayout
            //ExSummary:Shows how to update the layout of tables in a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            // Normally this method is not necessary to call, as cell and table widths are maintained automatically.
            // This method may need to be called when exporting to PDF in rare cases when the table layout appears
            // incorrectly in the rendered output.
            doc.UpdateTableLayout();
            //ExEnd
        }

        [Test]
        public void GetPageCount()
        {
            //ExStart
            //ExFor:Document.PageCount
            //ExSummary:Shows how to invoke page layout and retrieve the number of pages in the document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            // This invokes page layout which builds the document in memory so note that with large documents this
            // property can take time. After invoking this property, any rendering operation e.g rendering to PDF or image
            // will be instantaneous.
            int pageCount = doc.PageCount;
            //ExEnd

            Assert.AreEqual(1, pageCount);
        }

        [Test]
        public void UpdateFields()
        {
            //ExStart
            //ExFor:Document.UpdateFields
            //ExId:UpdateFieldsInDocument
            //ExSummary:Shows how to update all fields in a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            doc.UpdateFields();
            //ExEnd
        }

        [Test]
        public void GetUpdatedPageProperties()
        {
            //ExStart
            //ExFor:Document.UpdateWordCount()
            //ExFor:BuiltInDocumentProperties.Characters
            //ExFor:BuiltInDocumentProperties.Words
            //ExFor:BuiltInDocumentProperties.Paragraphs
            //ExSummary:Shows how to update all list labels in a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            // Some work should be done here that changes the document's content.

            // Update the word, character and paragraph count of the document.
            doc.UpdateWordCount();

            // Display the updated document properties.
            Console.WriteLine("Characters: {0}", doc.BuiltInDocumentProperties.Characters);
            Console.WriteLine("Words: {0}", doc.BuiltInDocumentProperties.Words);
            Console.WriteLine("Paragraphs: {0}", doc.BuiltInDocumentProperties.Paragraphs);
            //ExEnd
        }

        [Test]
        public void TableStyleToDirectFormatting()
        {
            //ExStart
            //ExFor:Document.ExpandTableStylesToDirectFormatting
            //ExId:TableStyleToDirectFormatting
            //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Table.TableStyle.docx");

            // Get the first cell of the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Cell firstCell = table.FirstRow.FirstCell;

            // First print the color of the cell shading. This should be empty as the current shading
            // is stored in the table style.
            Color cellShadingBefore = firstCell.CellFormat.Shading.BackgroundPatternColor;
            Console.WriteLine("Cell shading before style expansion: " + cellShadingBefore.ToString());

            // Expand table style formatting to direct formatting.
            doc.ExpandTableStylesToDirectFormatting();

            // Now print the cell shading after expanding table styles. A blue background pattern color
            // should have been applied from the table style.
            Color cellShadingAfter = firstCell.CellFormat.Shading.BackgroundPatternColor;
            Console.WriteLine("Cell shading after style expansion: " + cellShadingAfter.ToString());
            //ExEnd

            doc.Save(ExDir + "Table.ExpandTableStyleFormatting Out.docx");

            Assert.AreEqual(Color.Empty, cellShadingBefore);
            Assert.AreNotEqual(Color.Empty, cellShadingAfter);
        }

        [Test]
        public void GetOriginalFileInfo()
        {
            //ExStart
            //ExFor:Document.OriginalFileName
            //ExFor:Document.OriginalLoadFormat
            //ExSummary:Shows how to retrieve the details of the path, filename and LoadFormat of a document from when the document was first loaded into memory.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            // This property will return the full path and file name where the document was loaded from.
            string originalFilePath = doc.OriginalFileName;
            // Let's get just the file name from the full path.
            string originalFileName = Path.GetFileName(originalFilePath);
            
            // This is the original LoadFormat of the document.
            LoadFormat loadFormat = doc.OriginalLoadFormat;
            //ExEnd
        }

        [Test]
        public void RemoveSmartTagsFromDocument()
        {
            //ExStart
            //ExFor:CompositeNode.RemoveSmartTags
            //ExSummary:Shows how to remove all smart tags from a document.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            doc.RemoveSmartTags();
            //ExEnd
        }

        [Test]
        public void SetZoom()
        {
            //ExStart
            //ExFor:Document.ViewOptions
            //ExFor:ViewOptions
            //ExFor:ViewOptions.ViewType
            //ExFor:ViewOptions.ZoomPercent
            //ExFor:ViewType
            //ExId:SetZoom
            //ExSummary:The following code shows how to make sure the document is displayed at 50% zoom when opened in Microsoft Word.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            doc.ViewOptions.ViewType = ViewType.PageLayout;
            doc.ViewOptions.ZoomPercent = 50;
            doc.Save(ExDir + "Document.SetZoom Out.doc");
            //ExEnd
        }

        [Test]
        public void GetDocumentVariables()
        {
            //ExStart
            //ExFor:Document.Variables
            //ExFor:VariableCollection
            //ExId:GetDocumentVariables
            //ExSummary:Shows how to enumerate over document variables.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            foreach (DictionaryEntry entry in doc.Variables)
            {
                string name = entry.Key.ToString();
                string value = entry.Value.ToString();

                // Do something useful.
                Console.WriteLine("Name: {0}, Value: {1}", name, value);
            }
            //ExEnd
        }

        [Test]
        public void FootnoteOptionsEx()
        {
            //ExStart
            //ExFor:Document.FootnoteOptions
            //ExId:FootnoteOptionsEx
            //ExSummary:Shows how to edit a document's footnote options.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertFootnote(FootnoteType.Footnote, "My Footnote.");

            // Change your document's footnote options.
            doc.FootnoteOptions.Location = FootnoteLocation.BeneathText;
            doc.FootnoteOptions.NumberStyle = NumberStyle.Arabic;
            doc.FootnoteOptions.StartNumber = 1;

            doc.Save(ExDir + "Document.FootnoteOptions.doc");
            //ExEnd
        }

        [Test]
        public void HyphenationEx()
        {
            //ExStart
            //ExFor:Hyphenation
            //ExId:HyphenationEx
            //ExSummary:Load a hyphenation dictionary of a language from a file.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            
            Hyphenation.RegisterDictionary("en-US", ExDir + "hyph_en_US.dic");
            Console.Write(Hyphenation.IsDictionaryRegistered("en-US"));

            doc.Save(ExDir + "Document.HyphenationEx.doc");
            //ExEnd
        }

        [Test]
        public void CompareEx()
        {
            //ExStart
            //ExFor:Document.Compare
            //ExId:CompareEx
            //ExSummary:Shows how to apply the compare method to two documents and then use the results. 
            Aspose.Words.Document doc1 = new Aspose.Words.Document(ExDir + "Document.Compare.1.doc");
            Aspose.Words.Document doc2 = new Aspose.Words.Document(ExDir + "Document.Compare.2.doc");

            // Both documents should have no revisions or an exception will be thrown.
            if (doc1.Revisions.Count == 0 && doc2.Revisions.Count == 0)
                doc1.Compare(doc2, "authorName", DateTime.Now);

            // If doc1 and doc2 are different, doc1 now has some revisons after comparison, which can now be viewed and processed.
            foreach (Aspose.Words.Revision r in doc1.Revisions)
                Console.WriteLine(r.RevisionType);

            // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2.
            doc1.Revisions.AcceptAll();

            // doc1, when saved, now resembles doc2.
            doc1.Save(ExDir + "Document.CompareEx.doc");
            //ExEnd
        }

        [Test]
        public void RemoveExternalSchemaReferencesEx()
        {
            //ExStart
            //ExFor:Document.RemoveExternalSchemaReferences
            //ExId:RemoveExternalSchemaReferencesEx
            //ExSummary:Shows how to remove all external XML schema references from a document. 
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            doc.RemoveExternalSchemaReferences();
            //ExEnd
        }

        [Test]
        public void RemoveUnusedResourcesEx()
        {
            //ExStart
            //ExFor:Document.RemoveUnusedResources
            //ExId:RemoveUnusedResourcesEx
            //ExSummary:Shows how to remove all unused styles and lists from a document. 
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            doc.RemoveUnusedResources();
            //ExEnd
        }

        [Test]
        public void StartTrackRevisionsEx()
        {
            //ExStart
            //ExFor:Document.StartTrackRevisions
            //ExId:StartTrackRevisionsEx
            //ExSummary:Shows how StartTrackRevisions() affects document editing. 
            Aspose.Words.Document doc = new Aspose.Words.Document();
            doc.FirstSection.Body.FirstParagraph.Runs.Add(new Run(doc, "Hello world!"));

            Console.WriteLine(doc.Revisions.Count); // 0

            doc.StartTrackRevisions("author", DateTime.Now);

            doc.FirstSection.Body.AppendParagraph("Hello again!");

            Console.WriteLine(doc.Revisions.Count); // 2

            // The "Hello world!" text we added before doc.StartTrackRevisions() shows up as plain text in the output doc.
            // However, the "Hello again!" text we added after doc.StartTrackRevisions() is a revision in the output.
            doc.Save(ExDir + "Document.StartTrackRevisions.doc");
            //ExEnd
        }

        [Test]
        public void StopTrackRevisionsEx()
        {
            //ExStart
            //ExFor:Document.StopTrackRevisions
            //ExId:StopTrackRevisionsEx
            //ExSummary:Shows how to stop StartTrackRevisions(). 
            Aspose.Words.Document doc = new Aspose.Words.Document();
            doc.StopTrackRevisions();
            //ExEnd
        }

        [Test]
        public void UpdateThumbnailEx()
        {
            //ExStart
            //ExFor:Document.UpdateThumbnail
            //ExId:UpdateThumbnailEx
            //ExSummary:Shows how to update a document's thumbnail with and without ThumbnailGeneratingOptions.
            Aspose.Words.Document doc = new Aspose.Words.Document();

            // Update document's thumbnail the default way. 
            doc.UpdateThumbnail();

            // Review/change thumbnail options and then update document's thumbnail.
            Aspose.Words.Rendering.ThumbnailGeneratingOptions tgo
                = new Aspose.Words.Rendering.ThumbnailGeneratingOptions();

            Console.WriteLine("Thumbnail size: {0}", tgo.ThumbnailSize);
            tgo.GenerateFromFirstPage = true;

            doc.UpdateThumbnail(tgo);
            //ExEnd
        }
    }
}
