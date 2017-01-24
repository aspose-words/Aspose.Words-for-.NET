// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#if !JAVA
//ExStart
//ExId:ImportForDigitalSignatures
//ExSummary:The import required to use the X509Certificate2 class.

//ExEnd
#endif

using System;
using System.Collections;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Web;

using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Properties;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using Aspose.Words.Tables;
using Aspose.Words.Themes;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDocument : ApiExampleBase
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
            License license = new License();
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
                License license = new License();
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
            Document doc = new Document();
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
            Document doc = new Document(MyDir + "Document.doc");
            //ExEnd

            //ExStart
            //ExFor:Document.Save(String)
            //ExId:SaveToFile
            //ExSummary:Saves a document to a file.
            doc.Save(MyDir + @"\Artifacts\Document.OpenFromFile.doc");
            //ExEnd
        }

        [Test]
        public void OpenAndSaveToFile()
        {
            //ExStart
            //ExId:OpenAndSaveToFile
            //ExSummary:Opens a document from a file and saves it to a different format
            Document doc = new Document(MyDir + "Document.doc");
            doc.Save(MyDir + @"\Artifacts\Document.html");
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
            Stream stream = File.OpenRead(MyDir + "Document.doc");

            // Load the entire document into memory.
            Document doc = new Document(stream);

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
            string fileName = MyDir + "Document.OpenFromStreamWithBaseUri.html";

            // Open the stream.
            Stream stream = File.OpenRead(fileName);

            // Open the document. Note the Document constructor detects HTML format automatically.
            // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.BaseUri = MyDir;
            Document doc = new Document(stream, loadOptions);

            // You can close the stream now, it is no longer needed because the document is in memory.
            stream.Close();

            // Save in the DOC format.
            doc.Save(MyDir + @"\Artifacts\Document.OpenFromStreamWithBaseUri.doc");
            //ExEnd

            // Lets make sure the image was imported successfully into a Shape node.
            // Get the first shape node in the document.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            // Verify some properties of the image.
            Assert.IsTrue(shape.IsImage);
            Assert.IsNotNull(shape.ImageData.ImageBytes);
            Assert.AreEqual(80.0, ConvertUtil.PointToPixel(shape.Width));
            Assert.AreEqual(60.0, ConvertUtil.PointToPixel(shape.Height));
        }

        [Test]
        public void OpenDocumentFromWeb()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream)
            //ExSummary:Retrieves a document from a URL and saves it to disk in a different format.
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
            Document doc = new Document(byteStream);

            // Convert the document to any format supported by Aspose.Words.
            doc.Save(MyDir + @"\Artifacts\Document.OpenFromWeb.docx");
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
            Document doc = new Document(stream, options);

            // Save the document to disk.
            // The extension of the filename can be changed to save the document into other formats. e.g PDF, DOCX, ODT, RTF.
            doc.Save(MyDir + @"\Artifacts\Document.HtmlPageFromWebpage.doc");
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

            Document doc = new Document(MyDir + "Document.LoadFormat.html", loadOptions);
            //ExEnd
        }

        [Test]
        public void LoadFormatForOldDocuments()
        {
            //ExStart
            //ExFor:LoadFormat.DocPreWord60
            //ExSummary: Shows how to open older binary DOC format for Word6.0/Word95 documents
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFormat = Aspose.Words.LoadFormat.DocPreWord60;

            Document doc = new Document(MyDir + "Document.PreWord60.doc", loadOptions);
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
            Document doc = new Document(MyDir + "Document.LoadEncrypted.doc", new LoadOptions("qwerty"));
            //ExEnd
        }

        [Test]
        public void LoadEncryptedFromStream()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream,LoadOptions)
            //ExSummary:Loads a Microsoft Word document encrypted with a password from a stream.
            Stream stream = File.OpenRead(MyDir + "Document.LoadEncrypted.doc");
            Document doc = new Document(stream, new LoadOptions("qwerty"));
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
            Document doc = new Document(MyDir + "Document.doc");

            doc.Save(MyDir + @"\Artifacts\Document.ConvertToHtml.html", SaveFormat.Html);
            //ExEnd
        }

        [Test]
        public void ConvertToMhtml()
        {
            //ExStart
            //ExFor:Document.Save(String)
            //ExSummary:Converts from DOC to MHTML format.
            Document doc = new Document(MyDir + "Document.doc");

            doc.Save(MyDir + @"\Artifacts\Document.ConvertToMhtml.mht");
            //ExEnd
        }

        [Test]
        public void ConvertToTxt()
        {
            //ExStart
            //ExId:ExtractContentSaveAsText
            //ExSummary:Shows how to save a document in TXT format.
            Document doc = new Document(MyDir + "Document.doc");

            doc.Save(MyDir + @"\Artifacts\Document.ConvertToTxt.txt");
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
            Document doc = new Document(MyDir + "Document.doc");

            doc.Save(MyDir + @"\Artifacts\Document.Doc2PdfSave.pdf");
            //ExEnd
        }

        [Test]
        public void SaveToStream()
        {
            //ExStart
            //ExFor:Document.Save(Stream,SaveFormat)
            //ExId:SaveToStream
            //ExSummary:Shows how to save a document to a stream.
            Document doc = new Document(MyDir + "Document.doc");

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
            Document doc = new Document(MyDir + "Document.doc");

            doc.Save(Response, @"\Artifacts\Report.doc", ContentDisposition.Inline, null);
            //ExEnd
        }

        [Test]
        public void Doc2EpubSave()
        {
            //ExStart
            //ExId:Doc2EpubSave
            //ExSummary:Converts a document to EPUB using default save options.

            // Open an existing document from disk.
            Document doc = new Document(MyDir + "Document.EpubConversion.doc");

            // Save the document in EPUB format.
            doc.Save(MyDir + @"\Artifacts\Document.EpubConversion.epub");
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
            Document doc = new Document(MyDir + "Document.EpubConversion.doc");

            // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
            // how the output document is saved.
            HtmlSaveOptions saveOptions =
                new HtmlSaveOptions();

            // Specify the desired encoding.
            saveOptions.Encoding = Encoding.UTF8;

            // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
            // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
            // HTML files greater than a certain size e.g 300kb.
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

            // Specify that we want to export document properties.
            saveOptions.ExportDocumentProperties = true;

            // Specify that we want to save in EPUB format.
            saveOptions.SaveFormat = SaveFormat.Epub;

            // Export the document as an EPUB file.
            doc.Save(MyDir + @"\Artifacts\Document.EpubConversion.epub", saveOptions);
            //ExEnd
        }

        [Test]
        public void SaveHtmlPrettyFormat()
        {
            //ExStart
            //ExFor:SaveOptions.PrettyFormat
            //ExSummary:Shows how to pass an option to export HTML tags in a well spaced, human readable format.
            Document doc = new Document(MyDir + "Document.doc");

            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Enabling the PrettyFormat setting will export HTML in an indented format that is easy to read.
            // If this is setting is false (by default) then the HTML tags will be exported in condensed form with no indentation.
            htmlOptions.PrettyFormat = true;

            doc.Save(MyDir + @"\Artifacts\Document.PrettyFormat.html", htmlOptions);
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
            Document doc = new Document(MyDir + "Rendering.doc");

            // This is the directory we want the exported images to be saved to.
            string imagesDir = Path.Combine(MyDir, "Images");

            // The folder specified needs to exist and should be empty.
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportTextInputFormFieldAsText = true;
            options.ImagesFolder = imagesDir;

            doc.Save(MyDir + @"\Artifacts\Document.SaveWithOptions.html", options);
            //ExEnd

            // Verify the images were saved to the correct location.
            Assert.IsTrue(File.Exists(MyDir + @"\Artifacts\Document.SaveWithOptions.html"));
            Assert.AreEqual(9, Directory.GetFiles(imagesDir).Length);
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void SaveHtmlExportFontsCaller()
        {
            this.SaveHtmlExportFonts();
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
            Document doc = new Document(MyDir + "Document.doc");

            // Set the option to export font resources.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml);
            options.ExportFontResources = true;
            // Create and pass the object which implements the handler methods.
            options.FontSavingCallback = new HandleFontSaving();

            doc.Save(MyDir + @"\Artifacts\Document.SaveWithFontsExport.html", options);
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
            this.SaveHtmlExportImages();
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
            Document doc = new Document(MyDir + "Document.doc");

            // Create and pass the object which implements the handler methods.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ImageSavingCallback = new HandleImageSaving();

            doc.Save(MyDir + @"\Artifacts\Document.SaveWithCustomImagesExport.html", options);
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
            this.TestNodeChangingInDocument();
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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set up and pass the object which implements the handler methods.
            doc.NodeChangingCallback = new HandleNodeChangingFontChanger();

            // Insert sample HTML content
            builder.InsertHtml("<p>Hello World</p>");

            doc.Save(MyDir + @"\Artifacts\Document.FontChanger.doc");

            // Check that the inserted content has the correct formatting
            Run run = (Run)doc.GetChild(NodeType.Run, 0, true);
            Assert.AreEqual(24.0, run.Font.Size);
            Assert.AreEqual("Arial", run.Font.Name);
        }

        public class HandleNodeChangingFontChanger : INodeChangingCallback
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
            // Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format. These are both times where you might need extract the file format as it's not visible
            FileStream docStream = File.OpenRead(MyDir + "Document.FileWithoutExtension"); // The file format of this document is actually ".doc"
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
            Document doc = new Document(docStream);

            // Save the document with the original file name, " Out" and the document's file extension.
            doc.Save(MyDir + @"\Artifacts\Document.WithFileExtension" + FileFormatUtil.SaveFormatToExtension(saveFormat));
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
            Assert.AreEqual(".html", FileFormatUtil.LoadFormatToExtension(loadFormat));
        }

        [Test]
        public void AppendDocument()
        {
            //ExStart
            //ExFor:Document.AppendDocument(Document, ImportFormatMode)
            //ExSummary:Shows how to append a document to the end of another document.
            // The document that the content will be appended to.
            Document dstDoc = new Document(MyDir + "Document.doc");
            // The document to append.
            Document srcDoc = new Document(MyDir + "DocumentBuilder.doc");

            // Append the source document to the destination document.
            // Pass format mode to retain the original formatting of the source document when importing it.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the document.
            dstDoc.Save(MyDir + @"\Artifacts\Document.AppendDocument.doc");
            //ExEnd
        }

        [Test]
        // Using this file path keeps the example making sense when compared with automation so we expect
        // the file not to be found.
        public void AppendDocumentFromAutomation()
        {
            //ExStart
            //ExId:AppendDocumentFromAutomation
            //ExSummary:Shows how to join multiple documents together.
            // The document that the other documents will be appended to.
            Document doc = new Document();
            // We should call this method to clear this document of any existing content.
            doc.RemoveAllChildren();

            int recordCount = 5;
            for (int i = 1; i <= recordCount; i++)
            {
                Document srcDoc = new Document();
                
                // Open the document to join.
                Assert.That(() => srcDoc == new Document(@"C:\DetailsList.doc"), Throws.TypeOf<FileNotFoundException>());

                // Append the source document at the end of the destination document.
                doc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

                // In automation you were required to insert a new section break at this point, however in Aspose.Words we 
                // don't need to do anything here as the appended document is imported as separate sectons already.

                // If this is the second document or above being appended then unlink all headers footers in this section 
                // from the headers and footers of the previous section.
                if (i > 1)
                    Assert.That(() => doc.Sections[i].HeadersFooters.LinkToPrevious(false), Throws.TypeOf<NullReferenceException>());
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
            string filePath = MyDir + "Document.Signed.docx";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            if (info.HasDigitalSignature)
            {
                Console.WriteLine("Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.", Path.GetFileName(filePath));
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
            Document doc = new Document(MyDir + "Document.Signed.docx");

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
            Document doc = new Document(MyDir + "Document.Signed.docx");

            foreach (DigitalSignature signature in doc.DigitalSignatures)
            {
                Console.WriteLine("*** Signature Found ***");
                Console.WriteLine("Is valid: " + signature.IsValid);
                Console.WriteLine("Reason for signing: " + signature.Comments); // This property is available in MS Word documents only.
                Console.WriteLine("Signature type: " + signature.SignatureType.ToString());
                Console.WriteLine("Time of signing: " + signature.SignTime);
                Console.WriteLine("Subject name: " + signature.CertificateHolder.Certificate.SubjectName.ToString());
                Console.WriteLine("Issuer name: " + signature.CertificateHolder.Certificate.IssuerName.Name);
                Console.WriteLine();
            }
            //ExEnd

            DigitalSignature digitalSig = doc.DigitalSignatures[0];
            Assert.True(digitalSig.IsValid);
            Assert.AreEqual("Test Sign", digitalSig.Comments);
            Assert.AreEqual("XmlDsig", digitalSig.SignatureType.ToString());
            Assert.True(digitalSig.CertificateHolder.Certificate.Subject.Contains("Aspose Pty Ltd"));
            Assert.True(digitalSig.CertificateHolder.Certificate.IssuerName.Name != null && digitalSig.CertificateHolder.Certificate.IssuerName.Name.Contains("VeriSign"));
        }

        [Test]
        public void SignPdfDocument()
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:PdfDigitalSignatureDetails
            //ExFor:PdfSaveOptions.DigitalSignatureDetails
            //ExFor:PdfDigitalSignatureDetails.#ctor(X509Certificate2, String, String, DateTime)
            //ExId:SignPDFDocument
            //ExSummary:Shows how to sign a generated PDF document using Aspose.Words.
            // Create a simple document from scratch.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Test Signed PDF.");

            // Load the certificate from disk.
            // The other constructor overloads can be used to load certificates from different locations.
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            // Pass the certificate and details to the save options class to sign with.
            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(ch, "Test Signing", "Aspose Office", DateTime.Now);

            // Save the document as PDF with the digital signature set.
            doc.Save(MyDir + "Document.Signed Out.pdf", options);
            //ExEnd
        }

        //This is for obfuscation bug WORDSNET-13036
        [Test]
        public void SignDocument()
        {
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //By String
            Document doc = new Document(MyDir + "TestRepeatingSection.docx");
            string outputDocFileName = MyDir + @"\Artifacts\TestRepeatingSection.Signed.doc";

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "My comment", DateTime.Now);
        }

        [Test]
        public void AppendAllDocumentsInFolder()
        {
            string path = MyDir + @"\Artifacts\Document.AppendDocumentsFromFolder.doc";

            // Delete the file that was created by the previous run as I don't want to append it again.
            if (File.Exists(path))
                File.Delete(path);

            //ExStart
            //ExFor:Document.AppendDocument(Document, ImportFormatMode)
            //ExSummary:Shows how to use the AppendDocument method to combine all the documents in a folder to the end of a template document.
            // Lets start with a simple template and append all the documents in a folder to this document.
            Document baseDoc = new Document();

            // Add some content to the template.
            DocumentBuilder builder = new DocumentBuilder(baseDoc);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Template Document");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Some content here");

            // Gather the files which will be appended to our template document.
            // In this case we add the optional parameter to include the search only for files with the ".doc" extension.
            ArrayList files = new ArrayList(Directory.GetFiles(MyDir, "*.doc"));
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

                Document subDoc = new Document(fileName);
                baseDoc.AppendDocument(subDoc, ImportFormatMode.UseDestinationStyles);
            }

            // Save the combined document to disk.
            baseDoc.Save(path);
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
            Document doc = new Document(MyDir + "Rendering.doc");

            // This is for illustration purposes only, remember how many run nodes we had in the original document.
            int runsBefore = doc.GetChildNodes(NodeType.Run, true).Count;

            // Join runs with the same formatting. This is useful to speed up processing and may also reduce redundant
            // tags when exporting to HTML which will reduce the output file size.
            int joinCount = doc.JoinRunsWithSameFormatting();

            // This is for illustration purposes only, see how many runs are left after joining.
            int runsAfter = doc.GetChildNodes(NodeType.Run, true).Count;

            Console.WriteLine("Number of runs before:{0}, after:{1}, joined:{2}", runsBefore, runsAfter, joinCount);

            // Save the optimized document to disk.
            doc.Save(MyDir + @"\Artifacts\Document.JoinRunsWithSameFormatting.html");
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
            Document doc = new Document(MyDir + "Document.doc");
            doc.AttachedTemplate = "";
            doc.Save(MyDir + @"\Artifacts\Document.DetachTemplate.doc");
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
            Document doc = new Document(MyDir + "Document.doc");
            Document clone = doc.Clone();
            //ExEnd
        }

        [Test]
        public void ChangeFieldUpdateCultureSource()
        {
            // We will test this functionality creating a document with two fields with date formatting
            // field where the set language is different than the current culture, e.g German.
            Document doc = new Document();
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
            Document doc = new Document(MyDir + "Lists.PrintOutAllLists.doc");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // This option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
            // otherwise HTML <p> tag is used. This is also the default value.
            saveOptions.ExportListLabels = ExportListLabels.Auto;
            doc.Save(MyDir + @"\Artifacts\Document.ExportListLabels Auto.html", saveOptions);

            // Using this option the <p> tag is used for any list label representation.
            saveOptions.ExportListLabels = ExportListLabels.AsInlineText;
            doc.Save(MyDir + @"\Artifacts\Document.ExportListLabels InlineText.html", saveOptions);

            // The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
            saveOptions.ExportListLabels = ExportListLabels.ByHtmlTags;
            doc.Save(MyDir + @"\Artifacts\Document.ExportListLabels HtmlTags.html", saveOptions);
        }

        [Test]
        public void DocumentGetText_ToString()
        {
            //ExStart
            //ExFor:CompositeNode.GetText
            //ExFor:Node.ToString(SaveFormat)
            //ExId:NodeTxtExportDifferences
            //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
            Document doc = new Document();

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
            Document doc = new Document(MyDir + "Document.doc");

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
            Document loadDoc = new Document(inStream);
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
            Document doc = new Document();
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");
            //ExEnd

            //ExStart
            //ExFor:Document.Unprotect
            //ExId:UnprotectDocument
            //ExSummary:Shows how to unprotect a document. Note that the password is not required.
            doc.Unprotect();
            //ExEnd

            //ExStart
            //ExFor:Document.Unprotect(String)
            //ExSummary:Shows how to unprotect a document using a password.
            doc.Unprotect("password");
            //ExEnd
        }

        [Test]
        public void PasswordVerification()
        {
            Document doc = new Document();
            doc.WriteProtection.SetPassword("pwd");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.True(doc.WriteProtection.ValidatePassword("pwd"));
        }

        [Test]
        public void GetProtectionType()
        {
            //ExStart
            //ExFor:Document.ProtectionType
            //ExId:GetProtectionType
            //ExSummary:Shows how to get protection type currently set in the document.
            Document doc = new Document(MyDir + "Document.doc");
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
            Document doc = new Document();
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
            Document doc = new Document(MyDir + "Document.doc");
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
            Document doc = new Document(MyDir + "Document.doc");

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
            Document doc = new Document(MyDir + "Document.doc");

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
            Document doc = new Document(MyDir + "Document.doc");
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
            Document doc = new Document(MyDir + "Document.doc");

            // Some work should be done here that changes the document's content.

            // Update the word, character and paragraph count of the document.
            doc.UpdateWordCount();

            // Display the updated document properties.
            Console.WriteLine("Characters: {0}", doc.BuiltInDocumentProperties.Characters);
            Console.WriteLine("Words: {0}", doc.BuiltInDocumentProperties.Words);
            Console.WriteLine("Paragraphs: {0}", doc.BuiltInDocumentProperties.Paragraphs);
            //ExEnd
        }

        [Test, Explicit]
        public void TableStyleToDirectFormatting()
        {
            //ExStart
            //ExFor:Document.ExpandTableStylesToDirectFormatting
            //ExId:TableStyleToDirectFormatting
            //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
            Document doc = new Document(MyDir + "Table.TableStyle.docx");

            // Get the first cell of the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Cell firstCell = table.FirstRow.FirstCell;

            // First print the color of the cell shading. This should be empty as the current shading
            // is stored in the table style.
            double cellShadingBefore = table.FirstRow.RowFormat.Height;
            Console.WriteLine("Cell shading before style expansion: " + cellShadingBefore);

            // Expand table style formatting to direct formatting.
            doc.ExpandTableStylesToDirectFormatting();

            // Now print the cell shading after expanding table styles. A blue background pattern color
            // should have been applied from the table style.
            double cellShadingAfter = table.FirstRow.RowFormat.Height;
            Console.WriteLine("Cell shading after style expansion: " + cellShadingAfter);
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Table.ExpandTableStyleFormatting.docx");

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
            Document doc = new Document(MyDir + "Document.doc");

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
            Document doc = new Document(MyDir + "Document.doc");
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
            Document doc = new Document(MyDir + "Document.doc");
            doc.ViewOptions.ViewType = ViewType.PageLayout;
            doc.ViewOptions.ZoomPercent = 50;
            doc.Save(MyDir + @"\Artifacts\Document.SetZoom.doc");
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
            Document doc = new Document(MyDir + "Document.doc");

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
            //ExSummary:Shows how to insert a footnote and apply footnote options.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertFootnote(FootnoteType.Footnote, "My Footnote.");

            // Change your document's footnote options.
            doc.FootnoteOptions.Location = FootnoteLocation.BottomOfPage;
            doc.FootnoteOptions.NumberStyle = NumberStyle.Arabic;
            doc.FootnoteOptions.StartNumber = 1;

            doc.Save(MyDir + @"\Artifacts\Document.FootnoteOptions.doc");
            //ExEnd
        }

        [Test]
        public void CompareEx()
        {
            //ExStart
            //ExFor:Document.Compare
            //ExSummary:Shows how to apply the compare method to two documents and then use the results. 
            Document doc1 = new Document(MyDir + "Document.Compare.1.doc");
            Document doc2 = new Document(MyDir + "Document.Compare.2.doc");

            // If either document has a revision, an exception will be thrown.
            if (doc1.Revisions.Count == 0 && doc2.Revisions.Count == 0)
                doc1.Compare(doc2, "authorName", DateTime.Now);

            // If doc1 and doc2 are different, doc1 now has some revisons after the comparison, which can now be viewed and processed.
            foreach (Revision r in doc1.Revisions)
                Console.WriteLine(r.RevisionType);

            // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2.
            doc1.Revisions.AcceptAll();

            // doc1, when saved, now resembles doc2.
            doc1.Save(MyDir + @"\Artifacts\Document.CompareEx.doc");
            //ExEnd
        }

        [Test]
        public void CompareDocumentWithRevisions()
        {
            Document doc1 = new Document(MyDir + "Document.Compare.1.doc");
            Document docWithRevision = new Document(MyDir + "Document.Compare.Revisions.doc");
            
            if (docWithRevision.Revisions.Count > 0)
                Assert.That(() => docWithRevision.Compare(doc1, "authorName", DateTime.Now),
                Throws.TypeOf<InvalidOperationException>());
        }

        [Test]
        public void RemoveExternalSchemaReferencesEx()
        {
            //ExStart
            //ExFor:Document.RemoveExternalSchemaReferences
            //ExSummary:Shows how to remove all external XML schema references from a document. 
            Document doc = new Document(MyDir + "Document.doc");
            doc.RemoveExternalSchemaReferences();
            //ExEnd
        }

        [Test]
        public void RemoveUnusedResourcesEx()
        {
            //ExStart
            //ExFor:Document.RemoveUnusedResources
            //ExSummary:Shows how to remove all unused styles and lists from a document. 
            Document doc = new Document(MyDir + "Document.doc");
            doc.RemoveUnusedResources();
            //ExEnd
        }

        [Test]
        public void StartTrackRevisionsEx()
        {
            //ExStart
            //ExFor:Document.StartTrackRevisions(String)
            //ExFor:Document.StartTrackRevisions(String, DateTime)
            //ExFor:Document.StopTrackRevisions
            //ExSummary:Shows how tracking revisions affects document editing. 
            Document doc = new Document();

            // This text will appear as normal text in the document and no revisions will be counted.
            doc.FirstSection.Body.FirstParagraph.Runs.Add(new Run(doc, "Hello world!"));
            Console.WriteLine(doc.Revisions.Count); // 0

            doc.StartTrackRevisions("Author");

            // This text will appear as a revision. 
            // We did not specify a time while calling StartTrackRevisions(), so the date/time that's noted
            // on the revision will be the real time when StartTrackRevisions() executes.
            doc.FirstSection.Body.AppendParagraph("Hello again!");
            Console.WriteLine(doc.Revisions.Count); // 2

            // Stopping the tracking of revisions makes this text appear as normal text. 
            // Revisions are not counted when the document is changed.
            doc.StopTrackRevisions();
            doc.FirstSection.Body.AppendParagraph("Hello again!");
            Console.WriteLine(doc.Revisions.Count); // 2

            // Specifying some date/time will apply that date/time to all subsequent revisions until StopTrackRevisions() is called.
            // Note that placing values such as DateTime.MinValue as an argument will create revisions that do not have a date/time at all.
            doc.StartTrackRevisions("Author", new DateTime(1970, 1, 1));
            doc.FirstSection.Body.AppendParagraph("Hello again!");
            Console.WriteLine(doc.Revisions.Count); // 4

            doc.Save(MyDir + @"\Artifacts\Document.StartTrackRevisions.doc");
            //ExEnd
        }

        [Test]
        public void ShowRevisionBalloonsInPdf()
        {
            //ExStart
            //ExFor:RevisionOptions.ShowRevisionBalloons
            //ExSummary:Shows how render tracking changes in balloons
            Document doc = new Document(MyDir + "ShowRevisionBalloons.docx");

            //Set option true, if you need render tracking changes in balloons in pdf document
            doc.LayoutOptions.RevisionOptions.ShowRevisionBalloons = true;

            //Check that revisions are in balloons 
            doc.Save(MyDir + @"\Artifacts\ShowRevisionBalloons.pdf");
            //ExEnd
        }

        [Test]
        public void AcceptAllRevisions()
        {
            //ExStart
            //ExFor:Document.AcceptAllRevisions
            //ExSummary:Shows how to accept all tracking changes in the document.
            Document doc = new Document(MyDir + "Document.doc");

            // Start tracking and make some revisions.
            doc.StartTrackRevisions("Author");
            doc.FirstSection.Body.AppendParagraph("Hello world!");

            // Revisions will now show up as normal text in the output document.
            doc.AcceptAllRevisions();
            doc.Save(MyDir + @"\Artifacts\Document.AcceptedRevisions.doc");
            //ExEnd
        }

        [Test]
        public void UpdateThumbnailEx()
        {
            //ExStart
            //ExFor:Document.UpdateThumbnail()
            //ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
            //ExSummary:Shows how to update a document's thumbnail.
            Document doc = new Document();

            // Update document's thumbnail the default way. 
            doc.UpdateThumbnail();

            // Review/change thumbnail options and then update document's thumbnail.
            ThumbnailGeneratingOptions tgo = new ThumbnailGeneratingOptions();

            Console.WriteLine("Thumbnail size: {0}", tgo.ThumbnailSize);
            tgo.GenerateFromFirstPage = true;

            doc.UpdateThumbnail(tgo);
            //ExEnd
        }

        //For assert this test you need to open "HyphenationOptions OUT.docx" and check that hyphen are added in the end of the first line
        [Test]
        public void HyphenationOptions()
        {
            Document doc = new Document();

            DocumentHelper.InsertNewRun(doc, "poqwjopiqewhpefobiewfbiowefob ewpj weiweohiewobew ipo efoiewfihpewfpojpief pijewfoihewfihoewfphiewfpioihewfoihweoihewfpj", 0);

            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
            doc.HyphenationOptions.HyphenateCaps = true;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual(true, doc.HyphenationOptions.AutoHyphenation);
            Assert.AreEqual(2, doc.HyphenationOptions.ConsecutiveHyphenLimit);
            Assert.AreEqual(720, doc.HyphenationOptions.HyphenationZone);
            Assert.AreEqual(true, doc.HyphenationOptions.HyphenateCaps);

            doc.Save(MyDir + "HyphenationOptions.docx");
        }

        [Test]
        public void HyphenationOptionsDefaultValues()
        {
            Document doc = new Document();
            
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual(false, doc.HyphenationOptions.AutoHyphenation);
            Assert.AreEqual(0, doc.HyphenationOptions.ConsecutiveHyphenLimit);
            Assert.AreEqual(360, doc.HyphenationOptions.HyphenationZone); // 0.25 inch
            Assert.AreEqual(true, doc.HyphenationOptions.HyphenateCaps);
        }

        [Test]
        public void HyphenationOptionsExceptions()
        {
            Document doc = new Document();

            doc.HyphenationOptions.ConsecutiveHyphenLimit = 0;
            Assert.That(() => doc.HyphenationOptions.HyphenationZone = 0, Throws.TypeOf<ArgumentOutOfRangeException>());

            Assert.That(() => doc.HyphenationOptions.ConsecutiveHyphenLimit = -1, Throws.TypeOf<ArgumentOutOfRangeException>());
            doc.HyphenationOptions.HyphenationZone = 360;
        }

        [Ignore("Bug with .doc files")]
        [Test]
        public void ExtractPlainTextFromDocument()
        {
            //ExStart
            //ExFor:Document.ExtractText(string)
            //ExFor:Document.ExtractText(string, LoadOptions)
            //ExFor:PlaintextDocument.Text
            //ExFor:PlaintextDocument.BuiltInDocumentProperties
            //ExFor:PlaintextDocument.CustomDocumentProperties
            //ExSummary:Shows how to extract plain text from the document and get it properties
            PlainTextDocument plaintext = new PlainTextDocument(MyDir + "Bookmark.docx");
            Assert.AreEqual("This is a bookmarked text.\f", plaintext.Text); //in .doc there is other result "This is a bookmarked text.\r\r\r\r\r\r\r\f""

            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AllowTrailingWhitespaceForListItems = false;

            plaintext = new PlainTextDocument(MyDir + "Bookmark.doc", loadOptions);
            Assert.AreEqual("This is a bookmarked text.\f", plaintext.Text);

            BuiltInDocumentProperties builtInDocumentProperties = plaintext.BuiltInDocumentProperties;
            Assert.AreEqual("Aspose", builtInDocumentProperties.Company);

            CustomDocumentProperties customDocumentProperties = plaintext.CustomDocumentProperties;
            Assert.IsEmpty(customDocumentProperties);
            //ExEnd
        }

        [Ignore("Bug with .doc files")]
        [Test]
        public void ExtractPlainTextFromStream()
        {
            //ExStart
            //ExFor:Document.ExtractText(Stream)
            //ExFor:Document.ExtractText(Stream, LoadOptions)
            //ExSummary:
            Stream docStream = new FileStream(MyDir + "Bookmark.doc", FileMode.Open);

            PlainTextDocument plaintext = new PlainTextDocument(docStream);
            Assert.AreEqual("This is a bookmarked text.\f", plaintext.Text);

            docStream.Close();

            docStream = new FileStream(MyDir + "Bookmark.doc", FileMode.Open);

            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AllowTrailingWhitespaceForListItems = false;

            plaintext = new PlainTextDocument(docStream, loadOptions);
            Assert.AreEqual("This is a bookmarked text.\f", plaintext.Text);

            docStream.Close();
            //ExEnd
        }

        [Test]
        public void GetShapeAltTextTitle()
        {
            //ExStart
            //ExFor:Shape.Title
            //ExSummary:Shows how to get or set alt text title for shape object
            Document doc = new Document();

            // Create textbox shape.
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.Width = 431.5;
            shape.Height = 346.35;
            shape.Title = "Alt Text Title";

            Paragraph paragraph = new Paragraph(doc);
            paragraph.AppendChild(new Run(doc, "Test"));

            // Insert paragraph into the textbox.
            shape.AppendChild(paragraph);

            // Insert textbox into the document.
            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);
            
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Node[] shapes = doc.GetChildNodes(NodeType.Shape, true).ToArray();
            shape = (Shape)shapes[0];

            Assert.AreEqual("Alt Text Title", shape.Title);
            //ExEnd
        }

        [Test]
        public void GetOrSetDocumentThemeProperties()
        {
            Document doc = new Document();

            Theme theme = doc.Theme;

            theme.Colors.Accent1 = Color.Black;
            theme.Colors.Dark1 = Color.Blue;
            theme.Colors.FollowedHyperlink = Color.White;
            theme.Colors.Hyperlink = Color.WhiteSmoke;
            theme.Colors.Light1 = Color.Empty; //There is default Color.Black

            theme.MajorFonts.ComplexScript = "Arial";
            theme.MajorFonts.EastAsian = String.Empty;
            theme.MajorFonts.Latin = "Times New Roman";

            theme.MinorFonts.ComplexScript = String.Empty;
            theme.MinorFonts.EastAsian = "Times New Roman";
            theme.MinorFonts.Latin = "Arial";

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual(Color.Black.ToArgb(), doc.Theme.Colors.Accent1.ToArgb());
            Assert.AreEqual(Color.Blue.ToArgb(), doc.Theme.Colors.Dark1.ToArgb());
            Assert.AreEqual(Color.White.ToArgb(), doc.Theme.Colors.FollowedHyperlink.ToArgb());
            Assert.AreEqual(Color.WhiteSmoke.ToArgb(), doc.Theme.Colors.Hyperlink.ToArgb());
            Assert.AreEqual(Color.Black.ToArgb(), doc.Theme.Colors.Light1.ToArgb());

            Assert.AreEqual("Arial", doc.Theme.MajorFonts.ComplexScript);
            Assert.AreEqual(String.Empty, doc.Theme.MajorFonts.EastAsian);
            Assert.AreEqual("Times New Roman", doc.Theme.MajorFonts.Latin);

            Assert.AreEqual(String.Empty, doc.Theme.MinorFonts.ComplexScript);
            Assert.AreEqual("Times New Roman", doc.Theme.MinorFonts.EastAsian);
            Assert.AreEqual("Arial", doc.Theme.MinorFonts.Latin);
        }
    }
}