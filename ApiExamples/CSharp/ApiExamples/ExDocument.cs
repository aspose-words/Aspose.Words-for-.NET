// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Fonts;
using Aspose.Words.Layout;
using Aspose.Words.Lists;
using Aspose.Words.Loading;
using Aspose.Words.Markup;
using Aspose.Words.Properties;
using Aspose.Words.Rendering;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using Aspose.Words.Shaping.HarfBuzz;
using Aspose.Words.Tables;
using Aspose.Words.Themes;
using NUnit.Framework;
using CompareOptions = Aspose.Words.CompareOptions;
#if !(NETSTANDARD2_0 || __MOBILE__)
using Org.BouncyCastle.Pkcs;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDocument : ApiExampleBase
    {
#if !(__MOBILE__ || MAC)
        [Test]
        public void LicenseFromFileNoPath()
        {
            // This is where the test license is on my development machine.
            string testLicenseFileName = Path.Combine(LicenseDir, "Aspose.Words.lic");

            // Copy a license to the bin folder so the example can execute.
            string dstFileName = Path.Combine(AssemblyDir, "Aspose.Words.lic");
            File.Copy(testLicenseFileName, dstFileName);

            //ExStart
            //ExFor:License
            //ExFor:License.#ctor
            //ExFor:License.SetLicense(String)
            //ExId:LicenseFromFileNoPath
            //ExSummary:Aspose.Words will attempt to find the license file in the embedded resources or in the assembly folders.
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
            // This is where the test license is on my development machine.
            string testLicenseFileName = Path.Combine(LicenseDir, "Aspose.Words.lic");

            Stream myStream = File.OpenRead(testLicenseFileName);
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
#endif
        [Test]
        public void DocumentCtor()
        {
            //ExStart
            //ExId:DocumentCtor
            //ExFor:Document.#ctor(Boolean)
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
            doc.Save(ArtifactsDir + "Document.OpenFromFile.doc");
            //ExEnd
        }

        [Test]
        public void OpenAndSaveToFile()
        {
            //ExStart
            //ExId:OpenAndSaveToFile
            //ExSummary:Opens a document from a file and saves it to a different format
            Document doc = new Document(MyDir + "Document.doc");
            doc.Save(ArtifactsDir + "Document.html");
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
            using (Stream stream = File.OpenRead(MyDir + "Document.doc"))
            {
                // Load the entire document into memory.
                Document doc = new Document(stream);
                Assert.AreEqual("Hello World!\x000c", doc.GetText()); //ExSkip
            }
            // ... do something with the document
            //ExEnd
        }

        [Test]
        public void OpenFromStreamWithBaseUri()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream,LoadOptions)
            //ExFor:LoadOptions.#ctor
            //ExFor:LoadOptions.BaseUri
            //ExFor:ShapeBase.IsImage
            //ExId:DocumentCtor_LoadOptions
            //ExSummary:Opens an HTML document with images from a stream using a base URI.
            Document doc = new Document();
            // We are opening this HTML file:      
            //    <html>
            //    <body>
            //    <p>Simple file.</p>
            //    <p><img src="Aspose.Words.gif" width="80" height="60"></p>
            //    </body>
            //    </html>
            String fileName = MyDir + "Document.OpenFromStreamWithBaseUri.html";
            // Open the stream.
            using (Stream stream = File.OpenRead(fileName))
            {
                // Open the document. Note the Document constructor detects HTML format automatically.
                // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found.
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.BaseUri = MyDir;

                doc = new Document(stream, loadOptions);
            }

            // Save in the DOC format.
            doc.Save(ArtifactsDir + "Document.OpenFromStreamWithBaseUri.doc");
            //ExEnd

            // Lets make sure the image was imported successfully into a Shape node.
            // Get the first shape node in the document.
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
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
            String url = "https://is.gd/URJluZ";
            // The easiest way to load our document from the internet is make use of the 
            // System.Net.WebClient class. Create an instance of it and pass the URL
            // to download from.
            using (WebClient webClient = new WebClient())
            {
                // Download the bytes from the location referenced by the URL.
                byte[] dataBytes = webClient.DownloadData(url);

                // Wrap the bytes representing the document in memory into a MemoryStream object.
                using (MemoryStream byteStream = new MemoryStream(dataBytes))
                {
                    // Load this memory stream into a new Aspose.Words Document.
                    // The file format of the passed data is inferred from the content of the bytes itself. 
                    // You can load any document format supported by Aspose.Words in the same way.
                    Document doc = new Document(byteStream);

                    // Convert the document to any format supported by Aspose.Words.
                    doc.Save(ArtifactsDir + "Document.OpenFromWeb.docx");
                }
            }
            //ExEnd
        }

        [Test]
        public void InsertHtmlFromWebPage()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream, LoadOptions)
            //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
            //ExFor:LoadFormat
            //ExSummary:Shows how to insert the HTML contents from a web page into a new document.
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
            using (MemoryStream stream = new MemoryStream(pageBytes))
            {
                // The baseUri property should be set to ensure any relative img paths are retrieved correctly.
                LoadOptions options = new LoadOptions(Aspose.Words.LoadFormat.Html, "", url);

                // Load the HTML document from stream and pass the LoadOptions object.
                Document doc = new Document(stream, options);

                // Save the document to disk.
                // The extension of the filename can be changed to save the document into other formats. e.g PDF, DOCX, ODT, RTF.
                doc.Save(ArtifactsDir + "Document.HtmlPageFromWebpage.doc");
            }
            //ExEnd
        }

        [Test]
        public void LoadFormat()
        {
            //ExStart
            //ExFor:Document.#ctor(String,LoadOptions)
            //ExFor:LoadOptions.LoadFormat
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
            //ExFor:LoadFormat
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
            using (Stream stream = File.OpenRead(MyDir + "Document.LoadEncrypted.doc"))
            {
                Document doc = new Document(stream, new LoadOptions("qwerty"));
            }
            //ExEnd
        }

        [Test] 
        public void AnnotationsAtBlockLevel()
        {
            //ExStart
            //ExFor:LoadOptions.AnnotationsAtBlockLevel
            //ExFor:LoadOptions.AnnotationsAtBlockLevelAsDefault
            //ExSummary:Shows how to place bookmark nodes on the block, cell and row levels.
            // Any LoadOptions instances we create will have a default AnnotationsAtBlockLevel value equal to this
            LoadOptions.AnnotationsAtBlockLevelAsDefault = false;

            LoadOptions loadOptions = new LoadOptions();
            Assert.AreEqual(loadOptions.AnnotationsAtBlockLevel, LoadOptions.AnnotationsAtBlockLevelAsDefault);

            // If we want to work with annotations that transcend structures like tables, we will need to set this to true
            loadOptions.AnnotationsAtBlockLevel = true;

            // Open a document with a structured document tag and get that tag
            Document doc = new Document(MyDir + "Document.AnnotationsAtBlockLevel.docx", loadOptions);
            DocumentBuilder builder = new DocumentBuilder(doc);

            StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)[1];

            // Insert a bookmark and make it envelop our tag
            BookmarkStart start = builder.StartBookmark("MyBookmark");
            BookmarkEnd end = builder.EndBookmark("MyBookmark");

            sdt.ParentNode.InsertBefore(start, sdt);
            sdt.ParentNode.InsertAfter(end, sdt);

            doc.Save(ArtifactsDir + "Document.AnnotationsAtBlockLevel.docx", SaveFormat.Docx);
            //ExEnd
        }

        [Test]
        public void ConvertShapeToOfficeMath()
        {
            //ExStart
            //ExFor:LoadOptions.ConvertShapeToOfficeMath
            //ExSummary:Shows how to convert shapes with EquationXML to Office Math objects.
            LoadOptions loadOptions = new LoadOptions { ConvertShapeToOfficeMath = false };

            // Specify load option to convert math shapes to office math objects on loading stage.
            Document doc = new Document(MyDir + "Document.ConvertShapeToOfficeMath.docx", loadOptions);
            doc.Save(ArtifactsDir + "Document.ConvertShapeToOfficeMath.docx", SaveFormat.Docx);
            //ExEnd
        }

        [Test]
        public void LoadOptionsEncoding()
        {
            //ExStart
            //ExFor:LoadOptions.Encoding
            //ExSummary:Shows how to set the encoding with which to open a document.
            // Get the file format info of a file in our local file system
            FileFormatInfo fileFormatInfo = FileFormatUtil.DetectFileFormat(MyDir + "EncodedInUTF-7.txt");

            // One of the aspects of a document that the FileFormatUtil can pick up is the text encoding
            // This automatically takes place every time we open a document programmatically
            // Occasionally, due to the text content in the document as well as the lack of an encoding declaration,
            // the encoding of a document may be ambiguous 
            // In this case, while we know that our document is in UTF-7, the file encoding detector doesn't
            Assert.AreNotEqual(Encoding.UTF7, fileFormatInfo.Encoding);

            // If we open the document normally, the wrong encoding will be applied,
            // and the content of the document will not be represented correctly
            Document doc = new Document(MyDir + "EncodedInUTF-7.txt");
            Assert.AreEqual("Hello world+ACE-\r\n\r\n", doc.ToString(SaveFormat.Text));

            // In these cases we can set the Encoding attribute in a LoadOptions object
            // to override the automatically chosen encoding with the one we know to be correct
            LoadOptions loadOptions = new LoadOptions { Encoding = Encoding.UTF7 };
            doc = new Document(MyDir + "EncodedInUTF-7.txt", loadOptions);

            // This will give us the correct text
            Assert.AreEqual("Hello world!\r\n\r\n", doc.ToString(SaveFormat.Text));
            //ExEnd
        }

        [Test]
        public void LoadOptionsFontSettings()
        {
            //ExStart
            //ExFor:LoadOptions.FontSettings
            //ExSummary:Shows how to set font settings and apply them during the loading of a document. 
            // Create a FontSettings object that will substitute the "Times New Roman" font with the font "Arvo" from our "MyFonts" folder 
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolder(MyDir + @"MyFonts\", false);
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Times New Roman", "Arvo");

            // Set that FontSettings object as a member of a newly created LoadOptions object
            LoadOptions loadOptions = new LoadOptions { FontSettings = fontSettings };

            // We can now open a document while also passing the LoadOptions object into the constructor so the font substitution occurs upon loading
            Document doc = new Document(MyDir + "Document.docx", loadOptions);

            // The effects of our font settings can be observed after rendering
            doc.Save(ArtifactsDir + "Document.LoadOptionsFontSettings.pdf");
            //ExEnd
        }

        [Test]
        public void LoadOptionsMswVersion()
        {
            //ExStart
            //ExFor:LoadOptions.MswVersion
            //ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
            // Create a new LoadOptions object, which will load documents according to MS Word 2007 specification by default
            LoadOptions loadOptions = new LoadOptions();
            Assert.AreEqual(MsWordVersion.Word2007, loadOptions.MswVersion);

            // This document is missing the default paragraph format style,
            // so when it is opened with either Microsoft Word or Aspose Words, that default style will be regenerated,
            // and will show up in the Styles collection, with values according to Microsoft Word 2007 specifications
            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            Assert.AreEqual(13.8, doc.Styles.DefaultParagraphFormat.LineSpacing, 0.005f);

            // We can change the loading version like this, to Microsoft Word 2016
            loadOptions.MswVersion = MsWordVersion.Word2016;

            // The generated default style now has a different spacing, which will impact the appearance of our document
            doc = new Document(MyDir + "Document.docx", loadOptions);
            Assert.AreEqual(12.95, doc.Styles.DefaultParagraphFormat.LineSpacing, 0.005f);
            //ExEnd
        }

#if !(NETSTANDARD2_0 || __MOBILE__)
        //ExStart
        //ExFor:LoadOptions.ResourceLoadingCallback
        //ExSummary:Shows how to handle external resources in Html documents during loading.
        [Test] //ExSkip
        public void LoadOptionsCallback()
        {
            // Create a new LoadOptions object and set its ResourceLoadingCallback attribute
            // as an instance of our IResourceLoadingCallback implementation 
            LoadOptions loadOptions = new LoadOptions { ResourceLoadingCallback = new HtmlLinkedResourceLoadingCallback() };

            // When we open an Html document, external resources such as references to CSS stylesheet files and external images
            // will be handled in a custom manner by the loading callback as the document is loaded
            Document doc = new Document(MyDir + "ResourcesForCallback.html", loadOptions);
            doc.Save(ArtifactsDir + "Document.LoadOptionsCallback.pdf");
        }

        /// <summary>
        /// Resource loading callback that, upon encountering external resources,
        /// acknowledges CSS style sheets and replaces all images with a substitute.
        /// </summary>
        private class HtmlLinkedResourceLoadingCallback : IResourceLoadingCallback
        {
            public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
            {
                switch (args.ResourceType)
                {
                    case ResourceType.CssStyleSheet:
                        Console.WriteLine($"External CSS Stylesheet found upon loading: {args.OriginalUri}");
                        return ResourceLoadingAction.Default;
                    case ResourceType.Image:
                        Console.WriteLine($"External Image found upon loading: {args.OriginalUri}");

                        string newImageFilename =  "Images\\Aspose.Words.gif";
                        Console.WriteLine($"\tImage will be substituted with: {newImageFilename}");

                        System.Drawing.Image newImage = System.Drawing.Image.FromFile(MyDir + newImageFilename);

                        System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
                        byte[] imageBytes = (byte[])converter.ConvertTo(newImage, typeof(byte[]));
                        args.SetData(imageBytes);

                        return ResourceLoadingAction.UserProvided;

                }
                return ResourceLoadingAction.Default;
            }
        }
        //ExEnd
#endif

        //ExStart
        //ExFor:LoadOptions.WarningCallback
        //ExSummary:Shows how to print warnings that occur during document loading.
        [Test] //ExSkip
        public void LoadOptionsWarningCallback()
        {
            // Create a new LoadOptions object and set its WarningCallback attribute as an instance of our IWarningCallback implementation 
            LoadOptions loadOptions = new LoadOptions { WarningCallback = new DocumentLoadingWarningCallback() };

            // Minor warnings that might not prevent the effective loading of the document will now be printed
            Document doc = new Document(MyDir + "Document.docx", loadOptions);
        }

        /// <summary>
        /// IWarningCallback that prints warnings and their details as they arise during document loading.
        /// </summary>
        private class DocumentLoadingWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                Console.WriteLine($"WARNING: {info.WarningType}, source: {info.Source}");
                Console.WriteLine($"\tDescription: {info.Description}");
            }
        }
        //ExEnd

        [Test]
        public void ConvertToHtml()
        {
            //ExStart
            //ExFor:Document.Save(String,SaveFormat)
            //ExFor:SaveFormat
            //ExSummary:Converts from DOC to HTML format.
            Document doc = new Document(MyDir + "Document.doc");

            doc.Save(ArtifactsDir + "Document.ConvertToHtml.html", SaveFormat.Html);
            //ExEnd
        }

        [Test]
        public void ConvertToMhtml()
        {
            //ExStart
            //ExFor:Document.Save(String)
            //ExSummary:Converts from DOC to MHTML format.
            Document doc = new Document(MyDir + "Document.doc");

            doc.Save(ArtifactsDir + "Document.ConvertToMhtml.mht");
            //ExEnd
        }

        [Test]
        public void ConvertToTxt()
        {
            //ExStart
            //ExId:ExtractContentSaveAsText
            //ExSummary:Shows how to save a document in TXT format.
            Document doc = new Document(MyDir + "Document.doc");

            doc.Save(ArtifactsDir + "Document.ConvertToTxt.txt");
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

            doc.Save(ArtifactsDir + "Document.Doc2PdfSave.pdf");
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

            using (MemoryStream dstStream = new MemoryStream())
            {
                doc.Save(dstStream, SaveFormat.Docx);

                // Rewind the stream position back to zero so it is ready for next reader.
                dstStream.Position = 0;
            }
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
            doc.Save(ArtifactsDir + "Document.EpubConversion.epub");
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
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

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
            doc.Save(ArtifactsDir + "Document.EpubConversion.epub", saveOptions);
            //ExEnd
        }

        [Test]
        public void DownsampleOptions()
        {
            //ExStart
            //ExFor:DownsampleOptions
            //ExFor:DownsampleOptions.DownsampleImages
            //ExFor:DownsampleOptions.Resolution
            //ExFor:DownsampleOptions.ResolutionThreshold
            //ExSummary:Shows how to change the resolution of images in output pdf documents.
            // Open a document that contains images 
            Document doc = new Document(MyDir + "Rendering.doc");

            // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
            PdfSaveOptions options = new PdfSaveOptions();

            // This conversion will downsample images by default
            Assert.True(options.DownsampleOptions.DownsampleImages);
            Assert.AreEqual(220, options.DownsampleOptions.Resolution);

            // We can set the output resolution to a different value
            // The first two images in the input document will be affected by this
            options.DownsampleOptions.Resolution = 36;

            // We can set a minimum threshold for downsampling 
            // This value will prevent the second image in the input document from being downsampled
            options.DownsampleOptions.ResolutionThreshold = 128;

            doc.Save(ArtifactsDir + "PdfSaveOptions.DownsampleOptions.pdf", options);
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

            doc.Save(ArtifactsDir + "Document.PrettyFormat.html", htmlOptions);
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
            String imagesDir = Path.Combine(ArtifactsDir, "SaveHtmlWithOptions");

            // The folder specified needs to exist and should be empty.
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements.
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportTextInputFormFieldAsText = true;
            options.ImagesFolder = imagesDir;

            doc.Save(ArtifactsDir + "Document.SaveWithOptions.html", options);
            //ExEnd

            // Verify the images were saved to the correct location.
            Assert.IsTrue(File.Exists(ArtifactsDir + "Document.SaveWithOptions.html"));
            Assert.AreEqual(9, Directory.GetFiles(imagesDir).Length);

            Directory.Delete(imagesDir, true);
        }

        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontResources
        //ExFor:HtmlSaveOptions.FontSavingCallback
        //ExFor:IFontSavingCallback
        //ExFor:IFontSavingCallback.FontSaving
        //ExFor:FontSavingArgs
        //ExFor:FontSavingArgs.Bold
        //ExFor:FontSavingArgs.Document
        //ExFor:FontSavingArgs.FontFamilyName
        //ExFor:FontSavingArgs.FontFileName
        //ExFor:FontSavingArgs.FontStream
        //ExFor:FontSavingArgs.IsExportNeeded
        //ExFor:FontSavingArgs.IsSubsettingNeeded
        //ExFor:FontSavingArgs.Italic
        //ExFor:FontSavingArgs.KeepFontStreamOpen
        //ExFor:FontSavingArgs.OriginalFileName
        //ExFor:FontSavingArgs.OriginalFileSize
        //ExSummary:Shows how to define custom logic for handling font exporting when saving to HTML based formats.
        [Test] //ExSkip
        public void SaveHtmlExportFonts()
        {
            Document doc = new Document(MyDir + "Rendering.doc");

            // Set the option to export font resources
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportFontResources = true;
            // Create and pass the object which implements the handler methods
            options.FontSavingCallback = new HandleFontSaving();

            doc.Save(ArtifactsDir + "Document.SaveWithFontsExport.html", options);
        }

        /// <summary>
        /// Prints information about fonts and saves them alongside their output .html
        /// </summary>
        public class HandleFontSaving : IFontSavingCallback
        {
            void IFontSavingCallback.FontSaving(FontSavingArgs args)
            {
                // Print information about fonts
                Console.Write($"Font:\t{args.FontFamilyName}");
                if (args.Bold) Console.Write(", bold");
                if (args.Italic) Console.Write(", italic");
                Console.WriteLine($"\nSource:\t{args.OriginalFileName}, {args.OriginalFileSize} bytes\n");

                Assert.True(args.IsExportNeeded);
                Assert.True(args.IsSubsettingNeeded);

                // We can designate where each font will be saved by either specifying a file name, or creating a new stream
                args.FontFileName = args.OriginalFileName.Split('\\').Last();

                args.FontStream = 
                    new FileStream(ArtifactsDir + args.OriginalFileName.Split('\\').Last(), FileMode.Create);
                Assert.False(args.KeepFontStreamOpen);

                // We can access the source document from here also
                Assert.True(args.Document.OriginalFileName.EndsWith("Rendering.doc"));
            }
        }
        //ExEnd

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
        [Test] //ExSkip
        public void TestNodeChangingInDocument()
        {
            // Create a blank document object
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set up and pass the object which implements the handler methods.
            doc.NodeChangingCallback = new HandleNodeChangingFontChanger();

            // Insert sample HTML content
            builder.InsertHtml("<p>Hello World</p>");

            doc.Save(ArtifactsDir + "Document.FontChanger.doc");

            // Check that the inserted content has the correct formatting
            Run run = (Run) doc.GetChild(NodeType.Run, 0, true);
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
                    Aspose.Words.Font font = ((Run) args.Node).Font;
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
            dstDoc.Save(ArtifactsDir + "Document.AppendDocument.doc");
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
                Assert.That(() => srcDoc == new Document(@"C:\DetailsList.doc"),
                    Throws.TypeOf<FileNotFoundException>());

                // Append the source document at the end of the destination document.
                doc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

                // In automation you were required to insert a new section break at this point, however in Aspose.Words we 
                // don't need to do anything here as the appended document is imported as separate sections already.

                // If this is the second document or above being appended then unlink all headers footers in this section 
                // from the headers and footers of the previous section.
                if (i > 1)
                    Assert.That(() => doc.Sections[i].HeadersFooters.LinkToPrevious(false),
                        Throws.TypeOf<NullReferenceException>());
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
            //ExFor:DigitalSignatureCollection.Count
            //ExFor:DigitalSignatureCollection.Item(Int32)
            //ExFor:DigitalSignatureType
            //ExId:ValidateAllDocumentSignatures
            //ExSummary:Shows how to validate all signatures in a document.
            // Load the signed document.
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            DigitalSignatureCollection digitalSignatureCollection = doc.DigitalSignatures;

            if (digitalSignatureCollection.IsValid)
            {
                Console.WriteLine("Signatures belonging to this document are valid");
                Console.WriteLine(digitalSignatureCollection.Count);
                Console.WriteLine(digitalSignatureCollection[0].SignatureType);
            }
            else
            {
                Console.WriteLine("Signatures belonging to this document are NOT valid");
            }
            //ExEnd
        }

        [Test]
        [Ignore("WORDSXAND-132")]
        public void ValidateIndividualDocumentSignatures()
        {
            //ExStart
            //ExFor:CertificateHolder.Certificate
            //ExFor:Document.DigitalSignatures
            //ExFor:DigitalSignature
            //ExFor:DigitalSignatureCollection
            //ExFor:DigitalSignature.IsValid
            //ExFor:DigitalSignature.Comments
            //ExFor:DigitalSignature.SignTime
            //ExFor:DigitalSignature.SignatureType
            //ExFor:DigitalSignature.Certificate
            //ExId:ValidateIndividualSignatures
            //ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
            // Load the document which contains signature.
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");

            foreach (DigitalSignature signature in doc.DigitalSignatures)
            {
                Console.WriteLine("*** Signature Found ***");
                Console.WriteLine("Is valid: " + signature.IsValid);
                Console.WriteLine("Reason for signing: " +
                                  signature.Comments); // This property is available in MS Word documents only.
                Console.WriteLine("Signature type: " + signature.SignatureType);
                Console.WriteLine("Time of signing: " + signature.SignTime);
                Console.WriteLine("Subject name: " + signature.CertificateHolder.Certificate.SubjectName);
                Console.WriteLine("Issuer name: " + signature.CertificateHolder.Certificate.IssuerName.Name);
                Console.WriteLine();
            }
            //ExEnd

            DigitalSignature digitalSig = doc.DigitalSignatures[0];
            Assert.True(digitalSig.IsValid);
            Assert.AreEqual("Test Sign", digitalSig.Comments);
            Assert.AreEqual("XmlDsig", digitalSig.SignatureType.ToString());
            Assert.True(digitalSig.CertificateHolder.Certificate.Subject.Contains("Aspose Pty Ltd"));
            Assert.True(digitalSig.CertificateHolder.Certificate.IssuerName.Name != null &&
                        digitalSig.CertificateHolder.Certificate.IssuerName.Name.Contains("VeriSign"));
        }

        [Test]
        [Description("WORDSNET-16868")]
        public void SignPdfDocument()
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:PdfDigitalSignatureDetails
            //ExFor:PdfSaveOptions.DigitalSignatureDetails
            //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
            //ExId:SignPDFDocument
            //ExSummary:Shows how to sign a generated PDF document using Aspose.Words.
            // Create a simple document from scratch.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Test Signed PDF.");

            // Load the certificate from disk.
            // The other constructor overloads can be used to load certificates from different locations.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            // Pass the certificate and details to the save options class to sign with.
            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails =
                new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", DateTime.Now);

            // Save the document as PDF with the digital signature set.
            doc.Save(ArtifactsDir + "Document.Signed.pdf", options);
            //ExEnd
        }

#if !(NETSTANDARD2_0 || __MOBILE__)
        [Test]
        public void CertificateHolderCreate()
        {
            //ExStart
            //ExFor:CertificateHolder.Create(Byte[], SecureString)
            //ExFor:CertificateHolder.Create(Byte[], String)
            //ExFor:CertificateHolder.Create(String, String, String)
            //ExSummary:Shows how to create CertificateHolder objects.
            // 1: Load a PKCS #12 file into a byte array and apply its password to create the CertificateHolder
            byte[] certBytes = File.ReadAllBytes(MyDir + "morzal.pfx");
            CertificateHolder.Create(certBytes, "aw");

            // 2: Pass a SecureString which contains the password instead of a normal string
            SecureString password = new NetworkCredential("", "aw").SecurePassword;
            CertificateHolder.Create(certBytes, password);

            // 3: If the certificate has private keys corresponding to aliases, we can use the aliases to fetch their respective keys
            // First, we'll check for valid aliases like this
            using (FileStream certStream = new FileStream(MyDir + "morzal.pfx", FileMode.Open))
            {
                Pkcs12Store pkcs12Store = new Pkcs12Store(certStream, "aw".ToCharArray());
                IEnumerator enumerator = pkcs12Store.Aliases.GetEnumerator();

                while (enumerator.MoveNext())
                {
                    if (enumerator.Current != null)
                    {
                        string currentAlias = enumerator.Current.ToString();
                        if (pkcs12Store.IsKeyEntry(currentAlias) && pkcs12Store.GetKey(currentAlias).Key.IsPrivate)
                        {
                            Console.WriteLine($"Valid alias found: {enumerator.Current}");
                        }
                    }
                }
            }

            // For this file, we'll use an alias found above
            CertificateHolder.Create(MyDir + "morzal.pfx", "aw", "c20be521-11ea-4976-81ed-865fbbfc9f24");

            // If we leave the alias null, then the first possible alias that retrieves a private key will be used
            CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);
            //ExEnd
        }
#endif

        [Test]
        public void DigitalSignatureSign()
        {
            //ExStart
            //ExFor:DigitalSignature.CertificateHolder
            //ExFor:DigitalSignature.IssuerName
            //ExFor:DigitalSignature.SubjectName
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder)
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder)
            //ExSummary:Shows how to sign documents with X.509 certificates.
            // Open an unsigned document
            Document unSignedDoc = new Document(MyDir + "Document.docx");

            // Verify that it isn't signed
            Assert.False(FileFormatUtil.DetectFileFormat(MyDir + "Document.docx").HasDigitalSignature);
            Assert.AreEqual(0, unSignedDoc.DigitalSignatures.Count);

            // Create a CertificateHolder object from a PKCS #12 file, which we will use to sign the document
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);

            // There are 2 ways of saving a signed copy of a document to the local file system
            // 1: Designate unsigned input and signed output files by filename and sign with the passed CertificateHolder 
            DigitalSignatureUtil.Sign(MyDir + "Document.docx", ArtifactsDir + "Document.Signed.1.docx", 
                certificateHolder, new SignOptions() { SignTime = DateTime.Now } );

            // 2: Create a stream for the input file and one for the output and create a file, signed with the CertificateHolder, at the file system location determine
            using (FileStream inDoc = new FileStream(MyDir + "Document.docx", FileMode.Open))
            {
                using (FileStream outDoc = new FileStream(ArtifactsDir + "Document.Signed.2.docx", FileMode.Create))
                {
                    DigitalSignatureUtil.Sign(inDoc, outDoc, certificateHolder);
                }
            }

            // Verify that our documents are signed
            Document signedDoc = new Document(ArtifactsDir + "Document.Signed.1.docx");
            Assert.True(FileFormatUtil.DetectFileFormat(ArtifactsDir + "Document.Signed.1.docx").HasDigitalSignature);
            Assert.AreEqual(1,signedDoc.DigitalSignatures.Count);
            Assert.True(signedDoc.DigitalSignatures[0].IsValid);

            signedDoc = new Document(ArtifactsDir + "Document.Signed.2.docx");
            Assert.True(FileFormatUtil.DetectFileFormat(ArtifactsDir + "Document.Signed.2.docx").HasDigitalSignature);
            Assert.AreEqual(1, signedDoc.DigitalSignatures.Count);
            Assert.True(signedDoc.DigitalSignatures[0].IsValid);

            // These digital signatures will have some of the properties from the X.509 certificate from the .pfx file we used
            Assert.AreEqual("CN=Morzal.Me", signedDoc.DigitalSignatures[0].IssuerName);
            Assert.AreEqual("CN=Morzal.Me", signedDoc.DigitalSignatures[0].SubjectName);
            //ExEnd
        }

        [Test]
        public void AppendAllDocumentsInFolder()
        {
            String path = ArtifactsDir + "Document.AppendDocumentsFromFolder.doc";

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
            ArrayList files = new ArrayList(Directory.GetFiles(MyDir, "*.doc")
                .Where(file => file.EndsWith(".doc", StringComparison.CurrentCultureIgnoreCase)).ToArray());
            // The list of files may come in any order, let's sort the files by name so the documents are enumerated alphabetically.
            files.Sort();

            // Iterate through every file in the directory and append each one to the end of the template document.
            foreach (String fileName in files)
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
            doc.Save(ArtifactsDir + "Document.JoinRunsWithSameFormatting.html");
            //ExEnd

            // Verify that runs were joined in the document.
            Assert.That(runsAfter, Is.LessThan(runsBefore));
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
            doc.Save(ArtifactsDir + "Document.DetachTemplate.doc");
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
            doc.MailMerge.Execute(new[] { "Date1" }, new object[] { new DateTime(2011, 1, 01) });

            //ExStart
            //ExFor:Document.FieldOptions
            //ExFor:FieldOptions
            //ExFor:FieldOptions.FieldUpdateCultureSource
            //ExFor:FieldUpdateCultureSource
            //ExId:ChangeFieldUpdateCultureSource
            //ExSummary:Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
            // Set the culture used during field update to the culture used by the field.
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.MailMerge.Execute(new[] { "Date2" }, new object[] { new DateTime(2011, 1, 01) });
            //ExEnd

            // Verify the field update behavior is correct.
            Assert.AreEqual("Saturday, 1 January 2011 - Samstag, 1 Januar 2011", doc.Range.Text.Trim());

            // Restore the original culture.
            Thread.CurrentThread.CurrentCulture = currentCulture;
        }

        [Test]
        public void DocumentGetTextToString()
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
            MemoryStream streamOut = new MemoryStream();
            // Save the document to stream.
            doc.Save(streamOut, SaveFormat.Docx);

            // Convert the document to byte form.
            byte[] docBytes = streamOut.ToArray();

            // The bytes are now ready to be stored/transmitted.

            // Now reverse the steps to load the bytes back into a document object.
            MemoryStream streamIn = new MemoryStream(docBytes);

            // Load the stream into a new document object.
            Document loadDoc = new Document(streamIn);
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
            //ExStart
            //ExFor:WriteProtection.SetPassword(String)
            //ExSummary:Sets the write protection password for the document.
            Document doc = new Document();
            doc.WriteProtection.SetPassword("pwd");
            //ExEnd

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

        [Test]
        public void TableStyleToDirectFormatting()
        {
            //ExStart
            //ExFor:Document.ExpandTableStylesToDirectFormatting
            //ExId:TableStyleToDirectFormatting
            //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
            Document doc = new Document(MyDir + "Table.TableStyle.docx");

            // Get the first cell of the first table in the document.
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
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

            doc.Save(ArtifactsDir + "Table.ExpandTableStyleFormatting.docx");

            Assert.AreEqual(0.0d, cellShadingBefore);
            Assert.AreEqual(0.0d, cellShadingAfter);
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
            String originalFilePath = doc.OriginalFileName;
            // Let's get just the file name from the full path.
            String originalFileName = Path.GetFileName(originalFilePath);

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
            doc.Save(ArtifactsDir + "Document.SetZoom.doc");
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

            foreach (KeyValuePair<string, string> entry in doc.Variables)
            {
                String name = entry.Key;
                String value = entry.Value;

                // Do something useful.
                Console.WriteLine("Name: {0}, Value: {1}", name, value);
            }
            //ExEnd
        }

        [Test]
        [Description("WORDSNET-16099")]
        public void SetFootnoteNumberOfColumns()
        {
            //ExStart
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.Columns
            //ExSummary:Shows how to set the number of columns with which the footnotes area is formatted.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            Assert.AreEqual(0, doc.FootnoteOptions.Columns); //ExSkip

            // Lets change number of columns for footnotes on page. If columns value is 0 than footnotes area
            // is formatted with a number of columns based on the number of columns on the displayed page
            doc.FootnoteOptions.Columns = 2;
            doc.Save(ArtifactsDir + "Document.FootnoteOptions.docx");
            //ExEnd

            //Assert that number of columns gets correct
            doc = new Document(ArtifactsDir + "Document.FootnoteOptions.docx");
            Assert.AreEqual(2, doc.FirstSection.PageSetup.FootnoteOptions.Columns);
        }

        [Test]
        public void SetFootnotePosition()
        {
            //ExStart
            //ExFor:FootnoteOptions.Position
            //ExFor:FootnotePosition
            //ExSummary:Shows how to define footnote position in the document.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            doc.FootnoteOptions.Position = FootnotePosition.BeneathText;
            //ExEnd
        }

        [Test]
        public void SetFootnoteNumberFormat()
        {
            //ExStart
            //ExFor:FootnoteOptions.NumberStyle
            //ExSummary:Shows how to define numbering format for footnotes in the document.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            doc.FootnoteOptions.NumberStyle = NumberStyle.Arabic1;
            //ExEnd
        }

        [Test]
        public void SetFootnoteRestartNumbering()
        {
            //ExStart
            //ExFor:FootnoteOptions.RestartRule
            //ExFor:FootnoteNumberingRule
            //ExSummary:Shows how to define when automatic numbering for footnotes restarts in the document.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            doc.FootnoteOptions.RestartRule = FootnoteNumberingRule.RestartPage;
            //ExEnd
        }

        [Test]
        public void SetFootnoteStartingNumber()
        {
            //ExStart
            //ExFor:FootnoteOptions.StartNumber
            //ExSummary:Shows how to define the starting number or character for the first automatically numbered footnotes.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            doc.FootnoteOptions.StartNumber = 1;
            //ExEnd
        }

        [Test]
        public void SetEndnotePosition()
        {
            //ExStart
            //ExFor:EndnoteOptions
            //ExFor:EndnoteOptions.Position
            //ExFor:EndnotePosition
            //ExSummary:Shows how to define endnote position in the document.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            doc.EndnoteOptions.Position = EndnotePosition.EndOfSection;
            //ExEnd
        }

        [Test]
        public void SetEndnoteNumberFormat()
        {
            //ExStart
            //ExFor:EndnoteOptions.NumberStyle
            //ExSummary:Shows how to define numbering format for endnotes in the document.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            doc.EndnoteOptions.NumberStyle = NumberStyle.Arabic1;
            //ExEnd
        }

        [Test]
        public void SetEndnoteRestartNumbering()
        {
            //ExStart
            //ExFor:EndnoteOptions.RestartRule
            //ExSummary:Shows how to define when automatic numbering for endnotes restarts in the document.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            doc.EndnoteOptions.RestartRule = FootnoteNumberingRule.RestartPage;
            //ExEnd
        }

        [Test]
        public void SetEndnoteStartingNumber()
        {
            //ExStart
            //ExFor:EndnoteOptions.StartNumber
            //ExSummary:Shows how to define the starting number or character for the first automatically numbered endnotes.
            Document doc = new Document(MyDir + "Document.FootnoteEndnote.docx");

            doc.EndnoteOptions.StartNumber = 1;
            //ExEnd
        }

        [Test]
        public void CompareDocuments()
        {
            //ExStart
            //ExFor:Document.Compare(Document, String, DateTime)
            //ExFor:RevisionCollection.AcceptAll
            //ExSummary:Shows how to apply the compare method to two documents and then use the results. 
            Document doc1 = new Document(MyDir + "Document.Compare.1.doc");
            Document doc2 = new Document(MyDir + "Document.Compare.2.doc");

            // If either document has a revision, an exception will be thrown.
            if (doc1.Revisions.Count == 0 && doc2.Revisions.Count == 0)
                doc1.Compare(doc2, "authorName", DateTime.Now);

            // If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed.
            foreach (Revision r in doc1.Revisions)
                Console.WriteLine(r.RevisionType);

            // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2.
            doc1.Revisions.AcceptAll();

            // doc1, when saved, now resembles doc2.
            doc1.Save(ArtifactsDir + "Document.Compare.doc");
            //ExEnd
        }

        [Test]
        public void CompareDocumentsWithCompareOptions()
        {
            //ExStart
            //ExFor:CompareOptions
            //ExFor:CompareOptions.IgnoreFormatting
            //ExFor:CompareOptions.IgnoreCaseChanges
            //ExFor:CompareOptions.IgnoreComments
            //ExFor:CompareOptions.IgnoreTables
            //ExFor:CompareOptions.IgnoreFields
            //ExFor:CompareOptions.IgnoreFootnotes
            //ExFor:CompareOptions.IgnoreTextboxes
            //ExFor:CompareOptions.IgnoreHeadersAndFooters
            //ExFor:CompareOptions.Target
            //ExFor:ComparisonTargetType
            //ExFor:Document.Compare(Document, String, DateTime, CompareOptions)
            //ExSummary: Shows how to specify which document shall be used as a target during comparison
            Document doc1 = new Document(MyDir + "Document.CompareOptions.1.docx");
            Document doc2 = new Document(MyDir + "Document.CompareOptions.2.docx");

            // ComparisonTargetType with IgnoreFormatting setting determines which document has to be used as formatting source for ranges of equal text.
            CompareOptions compareOptions = new CompareOptions
            {
                IgnoreFormatting = true,
                IgnoreCaseChanges = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                Target = ComparisonTargetType.New
            };
            doc1.Compare(doc2, "vderyushev", DateTime.Now, compareOptions);

            doc1.Save(ArtifactsDir + "Document.CompareOptions.docx");
            //ExEnd
        }

        [Test]
        [Description("Result of this test is normal behavior MS Word. The bullet is missing for the 3rd list item")]
        public void UseCurrentDocumentFormattingWhenCompareDocuments()
        {
            Document doc1 = new Document(MyDir + "Document.CompareOptions.1.docx");
            Document doc2 = new Document(MyDir + "Document.CompareOptions.2.docx");

            Aspose.Words.CompareOptions compareOptions = new Aspose.Words.CompareOptions();
            compareOptions.IgnoreFormatting = true;
            compareOptions.Target = ComparisonTargetType.Current;

            doc1.Compare(doc2, "vderyushev", DateTime.Now, compareOptions);

            doc1.Save(ArtifactsDir + "Document.UseCurrentDocumentFormatting.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "Document.UseCurrentDocumentFormatting.docx",
                GoldsDir + "Document.UseCurrentDocumentFormatting Gold.docx"));
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
        public void RemoveExternalSchemaReferences()
        {
            //ExStart
            //ExFor:Document.RemoveExternalSchemaReferences
            //ExSummary:Shows how to remove all external XML schema references from a document. 
            Document doc = new Document(MyDir + "Document.doc");
            doc.RemoveExternalSchemaReferences();
            //ExEnd
        }

        [Test]
        public void RemoveUnusedResources()
        {
            //ExStart
            //ExFor:Document.Cleanup(CleanupOptions)
            //ExFor:CleanupOptions
            //ExFor:CleanupOptions.UnusedLists
            //ExFor:CleanupOptions.UnusedStyles
            //ExSummary:Shows how to remove all unused styles and lists from a document. 
            Document doc = new Document(MyDir + "Document.doc");
            
            CleanupOptions cleanupOptions = new CleanupOptions
            {
                UnusedLists = true,
                UnusedStyles = true
            };

            doc.Cleanup(cleanupOptions);
            //ExEnd
        }

        [Test]
        public void StartTrackRevisions()
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

            doc.Save(ArtifactsDir + "Document.StartTrackRevisions.doc");
            //ExEnd
        }

        [Test]
        public void ShowRevisionBalloonsInPdf()
        {
            //ExStart
            //ExFor:RevisionOptions.ShowInBalloons
            //ExSummary:Shows how render tracking changes in balloons
            Document doc = new Document(MyDir + "ShowRevisionBalloons.docx");

            //Set option true, if you need render tracking changes in balloons in pdf document
            doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.Format;

            //Check that revisions are in balloons 
            doc.Save(ArtifactsDir + "ShowRevisionBalloons.pdf");
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
            doc.Save(ArtifactsDir + "Document.AcceptedRevisions.doc");
            //ExEnd
        }

        [Test]
        public void RevisionHistory()
        {
            //ExStart
            //ExFor:Paragraph.IsMoveFromRevision
            //ExFor:Paragraph.IsMoveToRevision
            //ExFor:ParagraphCollection
            //ExFor:ParagraphCollection.Item(Int32)
            //ExSummary:Shows how to get paragraph that was moved (deleted/inserted) in Microsoft Word while change tracking was enabled.
            Document doc = new Document(MyDir + "Document.Revisions.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // There are two sets of move revisions in this document
            // One moves a small part of a paragraph, while the other moves a whole paragraph
            // Paragraph.IsMoveFromRevision/IsMoveToRevision will only be true if a whole paragraph is moved, as in the latter case
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].IsMoveFromRevision)
                    Console.WriteLine("The paragraph {0} has been moved (deleted).", i);
                if (paragraphs[i].IsMoveToRevision)
                    Console.WriteLine("The paragraph {0} has been moved (inserted).", i);
            }
            //ExEnd
        }

        [Test]
        public void UpdateThumbnail()
        {
            //ExStart
            //ExFor:Document.UpdateThumbnail()
            //ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
            //ExFor:ThumbnailGeneratingOptions
            //ExFor:ThumbnailGeneratingOptions.GenerateFromFirstPage
            //ExFor:ThumbnailGeneratingOptions.ThumbnailSize
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

        [Test]
        public void HyphenationOptions()
        {
            //ExStart
            //ExFor:Document.HyphenationOptions
            //ExFor:HyphenationOptions
            //ExFor:HyphenationOptions.AutoHyphenation
            //ExFor:HyphenationOptions.ConsecutiveHyphenLimit
            //ExFor:HyphenationOptions.HyphenationZone
            //ExFor:HyphenationOptions.HyphenateCaps
            //ExFor:ParagraphFormat.SuppressAutoHyphens
            //ExSummary:Shows how to configure document hyphenation options.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set this to insert a page break before this paragraph
            builder.Font.Size = 24;
            builder.ParagraphFormat.SuppressAutoHyphens = false;

            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
            doc.HyphenationOptions.HyphenateCaps = true;

            // Each paragraph has this flag that can be set to suppress hyphenation
            Assert.False(builder.ParagraphFormat.SuppressAutoHyphens);

            doc.Save(ArtifactsDir + "HyphenationOptions.docx");
            //ExEnd

            Assert.AreEqual(true, doc.HyphenationOptions.AutoHyphenation);
            Assert.AreEqual(2, doc.HyphenationOptions.ConsecutiveHyphenLimit);
            Assert.AreEqual(720, doc.HyphenationOptions.HyphenationZone);
            Assert.AreEqual(true, doc.HyphenationOptions.HyphenateCaps);

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "HyphenationOptions.docx",
                GoldsDir + "Document.HyphenationOptions Gold.docx"));
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

            Assert.That(() => doc.HyphenationOptions.ConsecutiveHyphenLimit = -1,
                Throws.TypeOf<ArgumentOutOfRangeException>());
            doc.HyphenationOptions.HyphenationZone = 360;
        }

        [Test]
        public void ExtractPlainTextFromDocument()
        {
            //ExStart
            //ExFor:PlainTextDocument
            //ExFor:PlainTextDocument.#ctor(String)
            //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
            //ExFor:PlainTextDocument.Text
            //ExSummary:Show how to simply extract text from a document.
            TxtLoadOptions loadOptions = new TxtLoadOptions { DetectNumberingWithWhitespaces = false };

            PlainTextDocument plaintext = new PlainTextDocument(MyDir + "Bookmark.docx");
            Assert.AreEqual("This is a bookmarked text.\f", plaintext.Text); //ExSkip 

            plaintext = new PlainTextDocument(MyDir + "Bookmark.docx", loadOptions);
            Assert.AreEqual("This is a bookmarked text.\f", plaintext.Text); //ExSkip
            //ExEnd
        }

        [Test]
        public void GetPlainTextBuiltInDocumentProperties()
        {
            //ExStart
            //ExFor:PlainTextDocument.BuiltInDocumentProperties
            //ExSummary:Show how to get BuiltIn properties of plain text document.
            PlainTextDocument plaintext = new PlainTextDocument(MyDir + "Bookmark.docx");
            BuiltInDocumentProperties builtInDocumentProperties = plaintext.BuiltInDocumentProperties;
            //ExEnd

            Assert.AreEqual("Aspose", builtInDocumentProperties.Company);
        }

        [Test]
        public void GetPlainTextCustomDocumentProperties()
        {
            //ExStart
            //ExFor:PlainTextDocument.CustomDocumentProperties
            //ExSummary:Show how to get custom properties of plain text document.
            PlainTextDocument plaintext = new PlainTextDocument(MyDir + "Bookmark.docx");
            CustomDocumentProperties customDocumentProperties = plaintext.CustomDocumentProperties;
            //ExEnd

            Assert.That(customDocumentProperties, Is.Empty);
        }

        [Test]
        public void ExtractPlainTextFromStream()
        {
            //ExStart
            //ExFor:PlainTextDocument.#ctor(Stream)
            //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
            //ExSummary:Show how to simply extract text from a stream.
            TxtLoadOptions loadOptions = new TxtLoadOptions { DetectNumberingWithWhitespaces = false };

            Stream stream = new FileStream(MyDir + "Bookmark.docx", FileMode.Open);

            PlainTextDocument plaintext = new PlainTextDocument(stream);
            Assert.AreEqual("This is a bookmarked text.\f", plaintext.Text); //ExSkip

            plaintext = new PlainTextDocument(stream, loadOptions);
            Assert.AreEqual("This is a bookmarked text.\f", plaintext.Text); //ExSkip
            //ExEnd

            stream.Close();
        }

        [Test]
        public void DocumentThemeProperties()
        {
            //ExStart
            //ExFor:Theme
            //ExFor:Theme.Colors
            //ExFor:Theme.MajorFonts
            //ExFor:Theme.MinorFonts
            //ExSummary:Show how to change document theme options.
            Document doc = new Document();
            // Get document theme and do something useful
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
            //ExEnd

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

        [Test]
        public void OoxmlComplianceVersion()
        {
            //ExStart
            //ExFor:Document.Compliance
            //ExSummary:Shows how to get OOXML compliance version.
            Document doc = new Document(MyDir + "Document.doc");

            OoxmlCompliance compliance = doc.Compliance;
            //ExEnd
            Assert.AreEqual(compliance, OoxmlCompliance.Ecma376_2006);

            doc = new Document(MyDir + "Field.BarCode.docx");
            compliance = doc.Compliance;

            Assert.AreEqual(compliance, OoxmlCompliance.Iso29500_2008_Transitional);
        }

        [Test]
        public void SaveWithOptions()
        {
            //ExStart
            //ExFor:Document.Save(Stream, String, Saving.SaveOptions)
            //ExSummary:Improve the quality of a rendered document with SaveOptions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Size = 60;

            builder.Writeln("Some text.");

            SaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

            options.UseAntiAliasing = false;
            doc.Save(ArtifactsDir + "Document.SaveOptionsLowQuality.jpg", options);

            options.UseAntiAliasing = true;
            options.UseHighQualityRendering = true;
            doc.Save(ArtifactsDir + "Document.SaveOptionsHighQuality.jpg", options);
            //ExEnd
        }

        [Test]
        public void WordCountUpdate()
        {
            //ExStart
            //ExFor:Document.UpdateWordCount(Boolean)
            //ExSummary:Shows how to keep track of the word count.
            // Create an empty document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is the first line.");
            builder.Writeln("This is the second line.");
            builder.Writeln("These three lines contain eighteen words in total.");

            // The fields that keep track of how many lines and words a document has are not automatically updated
            // An empty document has one paragraph by default, which contains one empty line
            Assert.AreEqual(0, doc.BuiltInDocumentProperties.Words);
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.Lines);

            // To update them we have to use this method
            // The default constructor updates just the word count
            doc.UpdateWordCount();

            Assert.AreEqual(18, doc.BuiltInDocumentProperties.Words);
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.Lines);

            // If we want to update the line count as well, we have to use this overload
            doc.UpdateWordCount(true);

            Assert.AreEqual(18, doc.BuiltInDocumentProperties.Words);
            Assert.AreEqual(3, doc.BuiltInDocumentProperties.Lines);
            //ExEnd
        }

        [Test]
        public void CleanUpStyles()
        {
            //ExStart
            //ExFor:Document.Cleanup
            //ExSummary:Shows how to remove unused styles and lists from a document.
            // Create a new document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Brand new documents have 4 styles and 0 lists by default
            Assert.AreEqual(4, doc.Styles.Count);
            Assert.AreEqual(0, doc.Lists.Count);

            // We will add one style and one list and mark them as "used" by applying them to the builder 
            builder.ParagraphFormat.Style = doc.Styles.Add(StyleType.Paragraph, "My Used Style");
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDiamonds);

            // These items were added to their respective collections
            Assert.AreEqual(5, doc.Styles.Count);
            Assert.AreEqual(1, doc.Lists.Count);

            // doc.Cleanup() removes all unused styles and lists
            doc.Cleanup();

            // It currently has no effect because the 2 items we added plus the original 4 styles are all used
            Assert.AreEqual(5, doc.Styles.Count);
            Assert.AreEqual(1, doc.Lists.Count);

            // These two items will be added but will not associated with any part of the document
            doc.Styles.Add(StyleType.Paragraph, "My Unused Style");
            doc.Lists.Add(ListTemplate.NumberArabicDot);

            // They also get stored in the document and are ready to be used
            Assert.AreEqual(6, doc.Styles.Count);
            Assert.AreEqual(2, doc.Lists.Count);

            doc.Cleanup();

            // Since we didn't apply them anywhere, the two unused items are removed by doc.Cleanup()
            Assert.AreEqual(5, doc.Styles.Count);
            Assert.AreEqual(1, doc.Lists.Count);
            //ExEnd
        }

        [Test]
        public void Revisions()
        {
            //ExStart
            //ExFor:Revision
            //ExFor:Revision.Accept
            //ExFor:Revision.Author
            //ExFor:Revision.DateTime
            //ExFor:Revision.Group
            //ExFor:Revision.Reject
            //ExFor:Revision.RevisionType
            //ExFor:RevisionCollection
            //ExFor:RevisionCollection.Item(Int32)
            //ExFor:RevisionCollection.Count
            //ExFor:Document.HasRevisions
            //ExFor:Document.TrackRevisions
            //ExFor:Document.Revisions
            //ExSummary:Shows how to check if a document has revisions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Normal editing of the document does not count as a revision
            builder.Write("This does not count as a revision. ");
            Assert.IsFalse(doc.HasRevisions);

            // In order for our edits to count as revisions, we need to declare an author and start tracking them
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            builder.Write("This is revision #1. ");

            // This flag corresponds to the "Track Changes" option being turned on in Microsoft Word, to track the editing manually
            // done there and not the programmatic changes we are about to do here
            Assert.IsFalse(doc.TrackRevisions);

            // As well as nodes in the document, revisions get referenced in this collection
            Assert.IsTrue(doc.HasRevisions);
            Assert.AreEqual(1, doc.Revisions.Count);

            Revision revision = doc.Revisions[0];
            Assert.AreEqual("John Doe", revision.Author);
            Assert.AreEqual("This is revision #1. ", revision.ParentNode.GetText());
            Assert.AreEqual(RevisionType.Insertion, revision.RevisionType);
            Assert.AreEqual(revision.DateTime.Date, DateTime.Now.Date);
            Assert.AreEqual(doc.Revisions.Groups[0], revision.Group);

            // Deleting content also counts as a revision
            // The most recent revisions are put at the start of the collection
            doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();
            Assert.AreEqual(RevisionType.Deletion, doc.Revisions[0].RevisionType);
            Assert.AreEqual(2, doc.Revisions.Count);

            // Insert revisions are treated as document text by the GetText() method before they are accepted,
            // since they are still nodes with text and are in the body
            Assert.AreEqual("This does not count as a revision. This is revision #1.", doc.GetText().Trim());

            // Accepting the deletion revision will assimilate it into the paragraph text and remove it from the collection
            doc.Revisions[0].Accept();
            Assert.AreEqual(1, doc.Revisions.Count);

            // Once the delete revision is accepted, the nodes that it concerns are removed and their text will not show up here
            Assert.AreEqual("This is revision #1.", doc.GetText().Trim());

            // The second insertion revision is now at index 0, which we can reject to ignore and discard it
            doc.Revisions[0].Reject();
            Assert.AreEqual(0, doc.Revisions.Count);
            Assert.AreEqual("", doc.GetText().Trim());

            // This takes us back to not counting changes as revisions
            doc.StopTrackRevisions();

            builder.Writeln("This also does not count as a revision.");
            Assert.AreEqual(0, doc.Revisions.Count);

            doc.Save(ArtifactsDir + "Document.Revisions.docx");
            //ExEnd
        }

        [Test]
        public void RevisionCollection()
        {
            //ExStart
            //ExFor:Revision.ParentStyle
            //ExFor:RevisionCollection.GetEnumerator
            //ExFor:RevisionCollection.Groups
            //ExFor:RevisionCollection.RejectAll
            //ExFor:RevisionGroupCollection.GetEnumerator
            //ExSummary:Shows how to look through a document's revisions.
            // Open a document that contains revisions and get its revision collection
            Document doc = new Document(MyDir + "Document.Revisions.docx");
            RevisionCollection revisions = doc.Revisions;
            
            // This collection itself has a collection of revision groups, which are merged sequences of adjacent revisions
            Console.WriteLine($"{revisions.Groups.Count} revision groups:");

            // We can iterate over the collection of groups and access the text that the revision concerns
            using (IEnumerator<RevisionGroup> e = revisions.Groups.GetEnumerator())
            {
                while (e.MoveNext())
                {
                    Console.WriteLine($"\tGroup type \"{e.Current.RevisionType}\", " +
                                      $"author: {e.Current.Author}, contents: [{e.Current.Text.Trim()}]");
                }
            }

            // The collection of revisions is considerably larger than the condensed form we printed above,
            // depending on how many Runs the text has been segmented into during editing in Microsoft Word,
            // since each Run affected by a revision gets its own Revision object
            Console.WriteLine($"\n{revisions.Count} revisions:");

            using (IEnumerator<Revision> e = revisions.GetEnumerator())
            {
                while (e.MoveNext())
                {
                    // A StyleDefinitionChange strictly affects styles and not document nodes, so in this case the ParentStyle
                    // attribute will always be used, while the ParentNode will always be null
                    // Since all other changes affect nodes, ParentNode will conversely be in use and ParentStyle will be null
                    if (e.Current.RevisionType == RevisionType.StyleDefinitionChange)
                    {
                        Console.WriteLine($"\tRevision type \"{e.Current.RevisionType}\", " +
                                          $"author: {e.Current.Author}, style: [{e.Current.ParentStyle.Name}]");
                    }
                    else
                    {
                        Console.WriteLine($"\tRevision type \"{e.Current.RevisionType}\", " +
                                          $"author: {e.Current.Author}, contents: [{e.Current.ParentNode.GetText().Trim()}]");
                    }
                }
            }

            // While the collection of revision groups provides a clearer overview of all revisions that took place in the document,
            // the changes must be accepted/rejected by the revisions themselves, the RevisionCollection, or the document
            // In this case we will reject all revisions via the collection, reverting the document to its original form, which we will then save
            revisions.RejectAll();
            Assert.AreEqual(0, revisions.Count);

            doc.Save(ArtifactsDir + "Document.RevisionCollection.docx");
            //ExEnd
        }

        [Test]
        public void AutoUpdateStyles()
        {
            //ExStart
            //ExFor:Document.AutomaticallyUpdateSyles
            //ExSummary:Shows how to update a document's styles based on its template.
            Document doc = new Document();

            // Empty Microsoft Word documents by default come with an attached template called "Normal.dotm"
            // There is no default template for Aspose Words documents
            Assert.AreEqual(string.Empty, doc.AttachedTemplate);

            // For AutomaticallyUpdateStyles to have any effect, we need a document with a template
            // We can make a document with word and open it
            // Or we can attach a template from our file system, as below
            doc.AttachedTemplate = MyDir + "Document.BusinessBrochureTemplate.dotx";

            Assert.IsTrue(doc.AttachedTemplate.EndsWith("Document.BusinessBrochureTemplate.dotx"));

            // Any changes to the styles in this template will be propagated to those styles in the document
            doc.AutomaticallyUpdateSyles = true;

            doc.Save(ArtifactsDir + "TemplateStylesUpdating.docx");
            //ExEnd
        }

        [Test]
        public void CompatibilityOptions()
        {
            //ExStart
            //ExFor:Document.CompatibilityOptions
            //ExSummary:Shows how to optimize our document for different word versions.
            Document doc = new Document();
            CompatibilityOptions co = doc.CompatibilityOptions;

            // Here are some default values
            Assert.AreEqual(true, co.GrowAutofit);
            Assert.AreEqual(false, co.DoNotBreakWrappedTables);
            Assert.AreEqual(false, co.DoNotUseEastAsianBreakRules);
            Assert.AreEqual(false, co.SelectFldWithFirstOrLastChar);
            Assert.AreEqual(false, co.UseWord97LineBreakRules);
            Assert.AreEqual(true, co.UseWord2002TableStyleRules);
            Assert.AreEqual(false, co.UseWord2010TableStyleRules);

            // This example covers only a small portion of all the compatibility attributes 
            // To see the entire list, in any of the output files go into File > Options > Advanced > Compatibility for...
            doc.Save(ArtifactsDir + "DefaultCompatibility.docx");

            // We can hand pick any value and change it to create a custom compatibility
            // We can also change a bunch of values at once to suit a defined compatibility scheme with the OptimizeFor method
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

            Assert.AreEqual(false, co.GrowAutofit);
            Assert.AreEqual(false, co.GrowAutofit);
            Assert.AreEqual(false, co.DoNotBreakWrappedTables);
            Assert.AreEqual(false, co.DoNotUseEastAsianBreakRules);
            Assert.AreEqual(false, co.SelectFldWithFirstOrLastChar);
            Assert.AreEqual(false, co.UseWord97LineBreakRules);
            Assert.AreEqual(false, co.UseWord2002TableStyleRules);
            Assert.AreEqual(true, co.UseWord2010TableStyleRules);

            doc.Save(ArtifactsDir + "Optimised for Word 2010.docx");

            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2000);

            Assert.AreEqual(true, co.GrowAutofit);
            Assert.AreEqual(true, co.DoNotBreakWrappedTables);
            Assert.AreEqual(true, co.DoNotUseEastAsianBreakRules);
            Assert.AreEqual(true, co.SelectFldWithFirstOrLastChar);
            Assert.AreEqual(false, co.UseWord97LineBreakRules);
            Assert.AreEqual(true, co.UseWord2002TableStyleRules);
            Assert.AreEqual(false, co.UseWord2010TableStyleRules);

            doc.Save(ArtifactsDir + "Optimised for Word 2000.docx");
            //ExEnd
        }

        [Test]
        public void Sections()
        {
            //ExStart
            //ExFor:Document.LastSection
            //ExSummary:Shows how to edit the last section of a document.
            // Open the template document, containing obsolete copyright information in the footer
            Document doc = new Document(MyDir + "HeaderFooter.ReplaceText.doc");

            // We have a document with 2 sections, this way FirstSection and LastSection are not the same
            Assert.AreEqual(2, doc.Sections.Count);

            string newCopyrightInformation = string.Format("Copyright (C) {0} by Aspose Pty Ltd.", DateTime.Now.Year);
            FindReplaceOptions findReplaceOptions =
                new FindReplaceOptions { MatchCase = false, FindWholeWordsOnly = false };

            // Access the first and the last sections
            HeaderFooter firstSectionFooter = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
            firstSectionFooter.Range.Replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

            HeaderFooter lastSectionFooter = doc.LastSection.HeadersFooters[HeaderFooterType.FooterPrimary];
            lastSectionFooter.Range.Replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

            // Sections are also accessible via an array
            Assert.AreEqual(doc.FirstSection, doc.Sections[0]);
            Assert.AreEqual(doc.LastSection, doc.Sections[1]);

            doc.Save(ArtifactsDir + "HeaderFooter.ReplaceText.doc");
            //ExEnd
        }

        [Test]
        public void DocTheme()
        {
            //ExStart
            //ExFor:Document.Theme
            //ExSummary:Shows what we can do with the Themes property of Document.
            Document doc = new Document();

            // When creating a blank document, Aspose Words creates a default theme object
            Theme theme = doc.Theme;

            // These color properties correspond to the 10 color columns that you see 
            // in the "Theme colors" section in the color selector menu when changing font or shading color
            // We can view and edit the leading color for each column, and the five different tints that
            // make up the rest of the column will be derived automatically from each leading color
            // Aspose Words sets the defaults to what they are in the Microsoft Word default theme
            Assert.AreEqual(Color.FromArgb(255, 255, 255, 255), theme.Colors.Light1);
            Assert.AreEqual(Color.FromArgb(255, 0, 0, 0), theme.Colors.Dark1);
            Assert.AreEqual(Color.FromArgb(255, 238, 236, 225), theme.Colors.Light2);
            Assert.AreEqual(Color.FromArgb(255, 31, 73, 125), theme.Colors.Dark2);
            Assert.AreEqual(Color.FromArgb(255, 79, 129, 189), theme.Colors.Accent1);
            Assert.AreEqual(Color.FromArgb(255, 192, 80, 77), theme.Colors.Accent2);
            Assert.AreEqual(Color.FromArgb(255, 155, 187, 89), theme.Colors.Accent3);
            Assert.AreEqual(Color.FromArgb(255, 128, 100, 162), theme.Colors.Accent4);
            Assert.AreEqual(Color.FromArgb(255, 75, 172, 198), theme.Colors.Accent5);
            Assert.AreEqual(Color.FromArgb(255, 247, 150, 70), theme.Colors.Accent6);

            // Hyperlink colors
            Assert.AreEqual(Color.FromArgb(255, 0, 0, 255), theme.Colors.Hyperlink);
            Assert.AreEqual(Color.FromArgb(255, 128, 0, 128), theme.Colors.FollowedHyperlink);

            // These appear at the very top of the font selector in the "Theme Fonts" section
            Assert.AreEqual("Cambria", theme.MajorFonts.Latin);
            Assert.AreEqual("Calibri", theme.MinorFonts.Latin);

            // Change some values to make a custom theme
            theme.MinorFonts.Latin = "Bodoni MT";
            theme.MajorFonts.Latin = "Tahoma";
            theme.Colors.Accent1 = Color.Cyan;
            theme.Colors.Accent2 = Color.Yellow;

            // Save the document to use our theme
            doc.Save(ArtifactsDir + "Document.Theme.docx");
            //ExEnd
        }

        [Test]
        public void SetEndnoteOptions()
        {
            //ExStart
            //ExFor:Document.EndnoteOptions
            //ExSummary:Shows how access a document's endnote options and see some of its default values.
            Document doc = new Document();

            Assert.AreEqual(1, doc.EndnoteOptions.StartNumber);
            Assert.AreEqual(EndnotePosition.EndOfDocument, doc.EndnoteOptions.Position);
            Assert.AreEqual(NumberStyle.LowercaseRoman, doc.EndnoteOptions.NumberStyle);
            Assert.AreEqual(FootnoteNumberingRule.Default, doc.EndnoteOptions.RestartRule);
            //ExEnd
        }

        [Test]
        public void SetInvalidateFieldTypes()
        {
            //ExStart
            //ExFor:Document.NormalizeFieldTypes
            //ExSummary:Shows how to get the keep a field's type up to date with its field code.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We'll add a date field
            Field field = builder.InsertField("DATE", null);

            // The FieldDate field type corresponds to the "DATE" field so our field's type property gets automatically set to it
            Assert.AreEqual(FieldType.FieldDate, field.Type);
            Assert.AreEqual(1, doc.Range.Fields.Count);

            // We can manually access the content of the field we added and change it
            Run fieldText = (Run) doc.FirstSection.Body.FirstParagraph.GetChildNodes(NodeType.Run, true)[0];
            Assert.AreEqual("DATE", fieldText.Text);
            fieldText.Text = "PAGE";

            // We changed the text to "PAGE" but the field's type property did not update accordingly
            Assert.AreEqual("PAGE", fieldText.GetText());
            Assert.AreNotEqual(FieldType.FieldPage, field.Type);

            // The type of the field as well as its components is still "FieldDate"
            Assert.AreEqual(FieldType.FieldDate, field.Type);
            Assert.AreEqual(FieldType.FieldDate, field.Start.FieldType);
            Assert.AreEqual(FieldType.FieldDate, field.Separator.FieldType);
            Assert.AreEqual(FieldType.FieldDate, field.End.FieldType);

            doc.NormalizeFieldTypes();

            // After running this method the type changes everywhere to "FieldPage", which matches the text "PAGE"
            Assert.AreEqual(FieldType.FieldPage, field.Type);
            Assert.AreEqual(FieldType.FieldPage, field.Start.FieldType);
            Assert.AreEqual(FieldType.FieldPage, field.Separator.FieldType);
            Assert.AreEqual(FieldType.FieldPage, field.End.FieldType);
            //ExEnd
        }

        [Test]
        public void DocLayoutOptions()
        {
            //ExStart
            //ExFor:Document.LayoutOptions
            //ExFor:LayoutOptions
            //ExFor:LayoutOptions.RevisionOptions
            //ExFor:Layout.LayoutOptions.ShowHiddenText
            //ExFor:Layout.LayoutOptions.ShowParagraphMarks
            //ExFor:RevisionColor
            //ExFor:RevisionOptions
            //ExFor:RevisionOptions.InsertedTextColor
            //ExFor:RevisionOptions.ShowRevisionBars
            //ExSummary:Shows how to set a document's layout options.
            Document doc = new Document();

            Assert.IsFalse(doc.LayoutOptions.ShowHiddenText);
            Assert.IsFalse(doc.LayoutOptions.ShowParagraphMarks);

            // The appearance of revisions can be controlled from the layout options property
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            doc.LayoutOptions.RevisionOptions.InsertedTextColor = RevisionColor.BrightGreen;
            doc.LayoutOptions.RevisionOptions.ShowRevisionBars = false;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln(
                "This is a revision. Normally the text is red with a bar to the left, but we made some changes to the revision options.");

            doc.StopTrackRevisions();

            // Layout options can be used to show hidden text too
            builder.Writeln("This text is not hidden.");
            builder.Font.Hidden = true;
            builder.Writeln(
                "This text is hidden. It will only show up in the output if we allow it to via doc.LayoutOptions.");

            doc.LayoutOptions.ShowHiddenText = true;

            doc.Save(ArtifactsDir + "Document.LayoutOptions.pdf");
            //ExEnd
        }

        [Test]
        public void DocMailMergeSettings()
        {
            //ExStart
            //ExFor:Document.MailMergeSettings
            //ExFor:MailMergeDataType
            //ExFor:MailMergeMainDocumentType
            //ExFor:Odso.Clone
            //ExFor:Odso.ColumnDelimiter
            //ExFor:Odso.DataSource
            //ExFor:Odso.DataSourceType
            //ExFor:Odso.FirstRowContainsColumnNames
            //ExSummary:Shows how to execute a mail merge with MailMergeSettings.
            // We'll create a simple document that will act as a destination for mail merge data
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(": ");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // Also we'll need a data source, in this case it will be an ASCII text file
            // We can use any character we want as a delimiter, in this case we'll choose '|'
            // The delimiter character is selected in the ODSO settings of mail merge settings
            string[] lines = { "FirstName|LastName|Message",
                "John|Doe|Hello! This message was created with Aspose Words mail merge." };
            File.WriteAllLines(ArtifactsDir + "Document.Lines.txt", lines);

            // Set the data source, query and other things
            MailMergeSettings mailMergeSettings = doc.MailMergeSettings;
            mailMergeSettings.MainDocumentType = MailMergeMainDocumentType.MailingLabels;
            mailMergeSettings.DataType = MailMergeDataType.Native;
            mailMergeSettings.DataSource = ArtifactsDir + "Document.Lines.txt";
            mailMergeSettings.Query = "SELECT * FROM " + doc.MailMergeSettings.DataSource;
            mailMergeSettings.LinkToQuery = true;
            mailMergeSettings.ViewMergedData = true;

            // Office Data Source Object settings
            Odso odso = mailMergeSettings.Odso;
            odso.DataSource = ArtifactsDir + "Document.Lines.txt";
            odso.DataSourceType = OdsoDataSourceType.Text;
            odso.ColumnDelimiter = '|';
            odso.FirstRowContainsColumnNames = true;

            // ODSO objects can also be cloned
            Assert.AreNotSame(odso, odso.Clone());

            // The mail merge will be performed when this document is opened 
            doc.Save(ArtifactsDir + "Document.MailMergeSettings.docx");
            //ExEnd
        }

        [Test]
        public void DocPackageCustomParts()
        {
            //ExStart
            //ExFor:CustomPart
            //ExFor:CustomPart.ContentType
            //ExFor:CustomPart.RelationshipType
            //ExFor:CustomPart.IsExternal
            //ExFor:CustomPart.Data
            //ExFor:CustomPart.Name
            //ExFor:CustomPart.Clone
            //ExFor:CustomPartCollection
            //ExFor:CustomPartCollection.Add(CustomPart)
            //ExFor:CustomPartCollection.Clear
            //ExFor:CustomPartCollection.Clone
            //ExFor:CustomPartCollection.Count
            //ExFor:CustomPartCollection.GetEnumerator
            //ExFor:CustomPartCollection.Item(Int32)
            //ExFor:CustomPartCollection.RemoveAt(Int32)
            //ExFor:Document.PackageCustomParts
            //ExSummary:Shows how to open a document with custom parts and access them.
            // Open a document that contains custom parts
            // CustomParts are arbitrary content OOXML parts
            // Not to be confused with Custom XML data which is represented by CustomXmlParts
            // This part is internal, meaning it is contained inside the OOXML package
            Document doc = new Document(MyDir + "Document.PackageCustomParts.docx");
            Assert.AreEqual(2, doc.PackageCustomParts.Count);

            // Clone the second part
            CustomPart clonedPart = doc.PackageCustomParts[1].Clone();

            // Add the clone to the collection
            doc.PackageCustomParts.Add(clonedPart);
            
            Assert.AreEqual(3, doc.PackageCustomParts.Count);

            // Use an enumerator to print information about the contents of each part 
            using (IEnumerator<CustomPart> enumerator = doc.PackageCustomParts.GetEnumerator())
            {
                int index = 0;
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"Part index {index}:");
                    Console.WriteLine($"\tName: {enumerator.Current.Name}");
                    Console.WriteLine($"\tContentType: {enumerator.Current.ContentType}");
                    Console.WriteLine($"\tRelationshipType: {enumerator.Current.RelationshipType}");
                    if (enumerator.Current.IsExternal)
                    {
                        Console.WriteLine("\tSourced from outside the document");
                    }
                    else
                    {
                        Console.WriteLine($"\tSourced from within the document, length: {enumerator.Current.Data.Length} bytes");
                    }
                    index++;
                }
            }

            TestCustomPartRead(doc); //ExSkip

            // Delete parts one at a time based on index
            doc.PackageCustomParts.RemoveAt(2);
            Assert.AreEqual(2, doc.PackageCustomParts.Count);

            // Delete all parts
            doc.PackageCustomParts.Clear();
            Assert.AreEqual(0, doc.PackageCustomParts.Count);
            //ExEnd
        }

        private void TestCustomPartRead(Document docWithCustomParts)
        {
            Assert.AreEqual("/payload/payload_on_package.test", docWithCustomParts.PackageCustomParts[0].Name); 
            Assert.AreEqual("mytest/somedata", docWithCustomParts.PackageCustomParts[0].ContentType); 
            Assert.AreEqual("http://mytest.payload.internal", docWithCustomParts.PackageCustomParts[0].RelationshipType); 
            Assert.AreEqual(false, docWithCustomParts.PackageCustomParts[0].IsExternal); 
            Assert.AreEqual(18, docWithCustomParts.PackageCustomParts[0].Data.Length); 

            // This part is external and its content is sourced from outside the document
            Assert.AreEqual("http://www.aspose.com/Images/aspose-logo.jpg", docWithCustomParts.PackageCustomParts[1].Name); 
            Assert.AreEqual("", docWithCustomParts.PackageCustomParts[1].ContentType); 
            Assert.AreEqual("http://mytest.payload.external", docWithCustomParts.PackageCustomParts[1].RelationshipType); 
            Assert.AreEqual(true, docWithCustomParts.PackageCustomParts[1].IsExternal); 
            Assert.AreEqual(0, docWithCustomParts.PackageCustomParts[1].Data.Length); 

            Assert.AreEqual("http://www.aspose.com/Images/aspose-logo.jpg", docWithCustomParts.PackageCustomParts[2].Name); 
            Assert.AreEqual("", docWithCustomParts.PackageCustomParts[2].ContentType); 
            Assert.AreEqual("http://mytest.payload.external", docWithCustomParts.PackageCustomParts[2].RelationshipType); 
            Assert.AreEqual(true, docWithCustomParts.PackageCustomParts[2].IsExternal); 
            Assert.AreEqual(0, docWithCustomParts.PackageCustomParts[2].Data.Length); 
        }

        [Test]
        public void DocShadeFormData()
        {
            //ExStart
            //ExFor:Document.ShadeFormData
            //ExSummary:Shows how to apply gray shading to bookmarks.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default, bookmarked text is highlighted gray
            Assert.IsTrue(doc.ShadeFormData);

            builder.Write("Text before bookmark. ");

            builder.InsertTextInput("My bookmark", TextFormFieldType.Regular, "",
                "If gray shading is turned on, this is the text that will have a gray background.", 0);

            // Our bookmarked text will appear gray here
            doc.Save(ArtifactsDir + "Document.ShadeFormDataTrue.docx");

            // In this file, shading will be turned off and the bookmarked text will blend in with the other text
            doc.ShadeFormData = false;
            doc.Save(ArtifactsDir + "Document.ShadeFormDataFalse.docx");
            //ExEnd
        }

        [Test]
        public void DocVersionsCount()
        {
            //ExStart
            //ExFor:Document.VersionsCount
            //ExSummary:Shows how to count how many previous versions a document has.
            Document doc = new Document();

            // No versions are in the document by default
            // We also can't add any since they are not supported
            Assert.AreEqual(0, doc.VersionsCount);

            // Let's open a document with versions
            doc = new Document(MyDir + "Versions.doc");

            // We can use this property to see how many there are
            Assert.AreEqual(4, doc.VersionsCount);

            doc.Save(ArtifactsDir + "Document.Versions.docx");      
            doc = new Document(ArtifactsDir + "Document.Versions.docx");

            // If we save and open the document, the versions are lost
            Assert.AreEqual(0, doc.VersionsCount);
            //ExEnd
        }

        [Test]
        public void DocWriteProtection()
        {
            //ExStart
            //ExFor:Document.WriteProtection
            //ExFor:WriteProtection
            //ExFor:WriteProtection.IsWriteProtected
            //ExFor:WriteProtection.ReadOnlyRecommended
            //ExFor:WriteProtection.ValidatePassword(String)
            //ExSummary:Shows how to protect a document with a password.
            Document doc = new Document();
            Assert.IsFalse(doc.WriteProtection.IsWriteProtected);
            Assert.IsFalse(doc.WriteProtection.ReadOnlyRecommended);

            // Enter a password that's 15 or less characters long
            doc.WriteProtection.SetPassword("docpassword123");
            Assert.IsTrue(doc.WriteProtection.IsWriteProtected);

            Assert.IsFalse(doc.WriteProtection.ValidatePassword("wrongpassword"));

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("We can still edit the document at this stage.");

            // Save the document
            // Without the password, we can only read this document in Microsoft Word
            // With the password, we can read and write
            doc.Save(ArtifactsDir + "Document.WriteProtection.docx");

            // Re-open our document
            Document docProtected = new Document(ArtifactsDir + "Document.WriteProtection.docx");
            DocumentBuilder docProtectedBuilder = new DocumentBuilder(docProtected);
            docProtectedBuilder.MoveToDocumentEnd();

            // We can programmatically edit this document without using our password
            Assert.IsTrue(docProtected.WriteProtection.IsWriteProtected);
            docProtectedBuilder.Writeln("Writing text in a protected document.");

            // We will still need the password if we want to open this one with Word
            docProtected.Save(ArtifactsDir + "Document.WriteProtectionEditedAfter.docx");
            //ExEnd
        }
        
        [Test]
        public void AddEditingLanguage()
        {
            //ExStart
            //ExFor:LanguagePreferences
            //ExFor:LanguagePreferences.AddEditingLanguage(EditingLanguage)
            //ExFor:LoadOptions.LanguagePreferences
            //ExFor:EditingLanguage
            //ExSummary:Shows how to set up language preferences that will be used when document is loading
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguagePreferences.AddEditingLanguage(EditingLanguage.Japanese);
            
            Document doc = new Document(MyDir + "Document.EditingLanguage.docx", loadOptions);

            int localeIdFarEast = doc.Styles.DefaultFont.LocaleIdFarEast;
            if (localeIdFarEast == (int)EditingLanguage.Japanese)
                Console.WriteLine("The document either has no any FarEast language set in defaults or it was set to Japanese originally.");
            else
                Console.WriteLine("The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
            //ExEnd
        }

        [Test]
        public void SetEditingLanguageAsDefault()
        {
            //ExStart
            //ExFor:LanguagePreferences.DefaultEditingLanguage
            //ExSummary:Shows how to set language as default
            LoadOptions loadOptions = new LoadOptions();
            // You can set language which only
            loadOptions.LanguagePreferences.DefaultEditingLanguage = EditingLanguage.Russian;

            Document doc = new Document(MyDir + "Document.EditingLanguage.docx", loadOptions);

            int localeId = doc.Styles.DefaultFont.LocaleId;
            if (localeId == (int)EditingLanguage.Russian)
                Console.WriteLine("The document either has no any language set in defaults or it was set to Russian originally.");
            else
                Console.WriteLine("The document default language was set to another than Russian language originally, so it is not overridden.");
            //ExEnd
        }

        [Test]
        public void GetInfoAboutRevisionsInRevisionGroups()
        {
            //ExStart
            //ExFor:RevisionGroup
            //ExFor:RevisionGroup.Author
            //ExFor:RevisionGroup.RevisionType
            //ExFor:RevisionGroup.Text
            //ExFor:RevisionGroupCollection
            //ExFor:RevisionGroupCollection.Count
            //ExSummary:Shows how to get info about a set of revisions in document.
            Document doc = new Document(MyDir + "Document.Revisions.docx");

            Console.WriteLine("Revision groups count: {0}\n", doc.Revisions.Groups.Count);

            // Get info about all of revisions in document
            foreach (RevisionGroup group in doc.Revisions.Groups)
            {
                Console.WriteLine("Revision author: {0}; Revision type: {1} \nRevision text: {2}", group.Author,
                    group.RevisionType, group.RevisionType);
            }

            //ExEnd
        }

        [Test]
        public void GetSpecificRevisionGroup()
        {
            //ExStart
            //ExFor:RevisionGroupCollection
            //ExFor:RevisionGroupCollection.Item(Int32)
            //ExFor:RevisionType
            //ExSummary:Shows how to get a set of revisions in document.
            Document doc = new Document(MyDir + "Document.Revisions.docx");

            // Get revision group by index.
            RevisionGroup revisionGroup = doc.Revisions.Groups[1];

            // Get info about specific revision groups sorted by RevisionType
            IEnumerable<string> revisionGroupCollectionInsertionType =
                doc.Revisions.Groups.Where(p => p.RevisionType == RevisionType.Insertion).Select(p =>
                    string.Format("Revision type: {0},\nRevision author: {1},\nRevision text: {2}.\n",
                        p.RevisionType.ToString(), p.Author, p.Text));

            foreach (string revisionGroupInfo in revisionGroupCollectionInsertionType)
            {
                Console.WriteLine(revisionGroupInfo);
            }
            //ExEnd
        }

        [Test]
        public void RemovePersonalInformation()
        {
            //ExStart
            //ExFor:Document.RemovePersonalInformation
            //ExSummary:Shows how to get or set a flag to remove all user information upon saving the MS Word document.
            Document doc = new Document(MyDir + "Document.docx")
            {
                // If flag sets to 'true' that MS Word will remove all user information from comments, revisions and
                // document properties upon saving the document. In MS Word 2013 and 2016 you can see this using
                // File -> Options -> Trust Center -> Trust Center Settings -> Privacy Options -> then the
                // checkbox "Remove personal information from file properties on save".
                RemovePersonalInformation = true
            };
            
            doc.Save(ArtifactsDir + "Document.RemovePersonalInformation.docx");
            //ExEnd
        }

        [Test]
        public void ShowComments()
        {
            //ExStart
            //ExFor:LayoutOptions.ShowComments
            //ExSummary:Shows how to show or hide comments in PDF document.
            Document doc = new Document(MyDir + "Comment.Document.docx");
            
            doc.LayoutOptions.ShowComments = false;
            
            doc.Save(ArtifactsDir + "Document.DoNotShowComments.pdf");
            //ExEnd
        }

        [Test]
        public void ShowRevisionsInBalloons()
        {
            //ExStart
            //ExFor:ShowInBalloons
            //ExFor:RevisionOptions.ShowInBalloons
            //ExFor:RevisionOptions.CommentColor
            //ExFor:RevisionOptions.DeletedTextColor
            //ExFor:RevisionOptions.DeletedTextEffect
            //ExFor:RevisionOptions.InsertedTextEffect
            //ExFor:RevisionOptions.MovedFromTextColor
            //ExFor:RevisionOptions.MovedFromTextEffect
            //ExFor:RevisionOptions.MovedToTextColor
            //ExFor:RevisionOptions.MovedToTextEffect
            //ExFor:RevisionOptions.RevisedPropertiesColor
            //ExFor:RevisionOptions.RevisedPropertiesEffect
            //ExFor:RevisionOptions.RevisionBarsColor
            //ExFor:RevisionOptions.RevisionBarsWidth
            //ExFor:RevisionOptions.ShowOriginalRevision
            //ExFor:RevisionOptions.ShowRevisionMarks
            //ExFor:RevisionTextEffect
            //ExSummary:Show how to render revisions in the balloons and edit their appearance.
            Document doc = new Document(MyDir + "Document.Revisions.docx");

            // Get the RevisionOptions object that controls the appearance of revisions
            RevisionOptions revisionOptions = doc.LayoutOptions.RevisionOptions;

            // Get movement, deletion, formatting revisions and comments to show up in green balloons on the right side of the page
            revisionOptions.ShowInBalloons = ShowInBalloons.Format;
            revisionOptions.CommentColor = RevisionColor.BrightGreen;

            // Render text inserted while revisions were being tracked in italic green
            revisionOptions.InsertedTextColor = RevisionColor.Green;
            revisionOptions.InsertedTextEffect = RevisionTextEffect.Italic;

            // Render text deleted while revisions were being tracked in bold red
            revisionOptions.DeletedTextColor = RevisionColor.Red;
            revisionOptions.DeletedTextEffect = RevisionTextEffect.Bold;

            // In a movement revision, the same text will appear twice: once at the departure point and once at the arrival destination
            // Render the text at the moved-from revision yellow with double strike through and double underlined blue at the moved-to revision
            revisionOptions.MovedFromTextColor = RevisionColor.Yellow;
            revisionOptions.MovedFromTextEffect = RevisionTextEffect.DoubleStrikeThrough;
            revisionOptions.MovedToTextColor = RevisionColor.Blue;
            revisionOptions.MovedFromTextEffect = RevisionTextEffect.DoubleUnderline;

            // Render text which had its format changed while revisions were being tracked in bold dark red
            revisionOptions.RevisedPropertiesColor = RevisionColor.DarkRed;
            revisionOptions.RevisedPropertiesEffect = RevisionTextEffect.Bold;

            // Place a thick dark blue bar on the left side of the page next to lines affected by revisions
            revisionOptions.RevisionBarsColor = RevisionColor.DarkBlue;
            revisionOptions.RevisionBarsWidth = 15.0f;

            // Show revision marks and original text
            revisionOptions.ShowOriginalRevision = true;
            revisionOptions.ShowRevisionMarks = true;

            doc.Save(ArtifactsDir + "Document.ShowRevisionsInBalloons.pdf");
            //ExEnd
        }

        [Test]
        public void CopyStylesFromTemplateViaDocument()
        {
            //ExStart
            //ExFor:Document.CopyStylesFromTemplate(Document)
            //ExSummary:Shows how to copies styles from the template to a document via Document.
            Document template = new Document(MyDir + "Rendering.doc");

            Document target = new Document(MyDir + "Document.docx");
            target.CopyStylesFromTemplate(template);

            target.Save(ArtifactsDir + "CopyStylesFromTemplateViaDocument.docx");
            //ExEnd
        }

        [Test]
        public void CopyStylesFromTemplateViaString()
        {
            //ExStart
            //ExFor:Document.CopyStylesFromTemplate(String)
            //ExSummary:Shows how to copies styles from the template to a document via string.
            string templatePath = MyDir + "Rendering.doc";
            
            Document target = new Document(MyDir + "Document.docx");
            target.CopyStylesFromTemplate(templatePath);

            target.Save(ArtifactsDir + "CopyStylesFromTemplateViaString.docx");
            //ExEnd
        }

        [Test]
        public void LayoutCollector()
        {
            //ExStart
            //ExFor:Layout.LayoutCollector
            //ExFor:Layout.LayoutCollector.#ctor(Document)
            //ExFor:Layout.LayoutCollector.Clear
            //ExFor:Layout.LayoutCollector.Document
            //ExFor:Layout.LayoutCollector.GetEndPageIndex(Node)
            //ExFor:Layout.LayoutCollector.GetEntity(Node)
            //ExFor:Layout.LayoutCollector.GetNumPagesSpanned(Node)
            //ExFor:Layout.LayoutCollector.GetStartPageIndex(Node)
            //ExFor:Layout.LayoutEnumerator.Current
            //ExSummary:Shows how to see the page spans of nodes.
            // Open a blank document and create a DocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a LayoutCollector object for our document that will have information about the nodes we placed
            LayoutCollector layoutCollector = new LayoutCollector(doc);

            // The document itself is a node that contains everything, which currently spans 0 pages
            Assert.AreEqual(doc, layoutCollector.Document);
            Assert.AreEqual(0, layoutCollector.GetNumPagesSpanned(doc));

            // Populate the document with sections and page breaks
            builder.Write("Section 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);
            doc.AppendChild(new Section(doc));
            doc.LastSection.AppendChild(new Body(doc));
            builder.MoveToDocumentEnd();
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);

            // The collected layout data won't automatically keep up with the real document contents
            Assert.AreEqual(0, layoutCollector.GetNumPagesSpanned(doc));

            // After we clear the layout collection and update it, the layout entity collection will be populated with up-to-date information about our nodes
            // The page span for the document now shows 5, which is what we would expect after placing 4 page breaks
            layoutCollector.Clear();
            doc.UpdatePageLayout();
            Assert.AreEqual(5, layoutCollector.GetNumPagesSpanned(doc));

            // We can also see the start/end pages of any other node, and their overall page spans
            NodeCollection nodes = doc.GetChildNodes(NodeType.Any, true);
            foreach (Node node in nodes)
            {
                Console.WriteLine($"->  NodeType.{node.NodeType}: ");
                Console.WriteLine($"\tStarts on page {layoutCollector.GetStartPageIndex(node)}, ends on page {layoutCollector.GetEndPageIndex(node)}, spanning {layoutCollector.GetNumPagesSpanned(node)} pages.");
            }

            // We can iterate over the layout entities using a LayoutEnumerator
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
            Assert.AreEqual(LayoutEntityType.Page, layoutEnumerator.Type);

            // The LayoutEnumerator can traverse the collection of layout entities like a tree
            // We can also point it to any node's corresponding layout entity like this
            layoutEnumerator.Current = layoutCollector.GetEntity(doc.GetChild(NodeType.Paragraph, 1, true));
            Assert.AreEqual(LayoutEntityType.Span, layoutEnumerator.Type);
            Assert.AreEqual("¶", layoutEnumerator.Text);
            //ExEnd
        }

        //ExStart
        //ExFor:Layout.LayoutEntityType
        //ExFor:Layout.LayoutEnumerator
        //ExFor:Layout.LayoutEnumerator.#ctor(Document)
        //ExFor:Layout.LayoutEnumerator.Document
        //ExFor:Layout.LayoutEnumerator.Kind
        //ExFor:Layout.LayoutEnumerator.MoveFirstChild
        //ExFor:Layout.LayoutEnumerator.MoveLastChild
        //ExFor:Layout.LayoutEnumerator.MoveNext
        //ExFor:Layout.LayoutEnumerator.MoveNextLogical
        //ExFor:Layout.LayoutEnumerator.MoveParent
        //ExFor:Layout.LayoutEnumerator.MoveParent(Layout.LayoutEntityType)
        //ExFor:Layout.LayoutEnumerator.MovePrevious
        //ExFor:Layout.LayoutEnumerator.MovePreviousLogical
        //ExFor:Layout.LayoutEnumerator.PageIndex
        //ExFor:Layout.LayoutEnumerator.Rectangle
        //ExFor:Layout.LayoutEnumerator.Reset
        //ExFor:Layout.LayoutEnumerator.Text
        //ExFor:Layout.LayoutEnumerator.Type
        //ExSummary:Demonstrates ways of traversing a document's layout entities.
        [Test] //ExSkip
        public void LayoutEnumerator()
        {
            // Open a document that contains a variety of layout entities
            // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
            // They are defined visually by the rectangular space that they occupy in the document
            Document doc = new Document(MyDir + "Document.LayoutEntities.docx");

            // Create an enumerator that can traverse these entities
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
            Assert.AreEqual(doc, layoutEnumerator.Document);

            // The enumerator points to the first element on the first page and can be traversed like a tree
            layoutEnumerator.MoveFirstChild();
            layoutEnumerator.MoveFirstChild();
            layoutEnumerator.MoveLastChild();
            layoutEnumerator.MovePrevious();
            Assert.AreEqual(LayoutEntityType.Span, layoutEnumerator.Type);
            Assert.AreEqual("TTT", layoutEnumerator.Text);

            // Only spans can contain text
            layoutEnumerator.MoveParent(LayoutEntityType.Page);
            Assert.AreEqual(LayoutEntityType.Page, layoutEnumerator.Type);

            // We can call this method to make sure that the enumerator points to the very first entity before we go through it forwards
            layoutEnumerator.Reset();

            // "Visual order" means when moving through an entity's children that are broken across pages,
            // page layout takes precedence and we avoid elements in other pages and move to others on the same page
            Console.WriteLine("Traversing from first to last, elements between pages separated:");
            TraverseLayoutForward(layoutEnumerator, 1);

            // Our enumerator is conveniently at the end of the collection for us to go through the collection backwards
            Console.WriteLine("Traversing from last to first, elements between pages separated:");
            TraverseLayoutBackward(layoutEnumerator, 1);

            // "Logical order" means when moving through an entity's children that are broken across pages, 
            // node relationships take precedence
            Console.WriteLine("Traversing from first to last, elements between pages mixed:");
            TraverseLayoutForwardLogical(layoutEnumerator, 1);

            Console.WriteLine("Traversing from last to first, elements between pages mixed:");
            TraverseLayoutBackwardLogical(layoutEnumerator, 1);
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order
        /// </summary>
        private void TraverseLayoutForward(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveFirstChild())
                {
                    TraverseLayoutForward(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MoveNext());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order
        /// </summary>
        private void TraverseLayoutBackward(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveLastChild())
                {
                    TraverseLayoutBackward(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MovePrevious());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order
        /// </summary>
        private void TraverseLayoutForwardLogical(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveFirstChild())
                {
                    TraverseLayoutForwardLogical(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MoveNextLogical());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order
        /// </summary>
        private void TraverseLayoutBackwardLogical(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveLastChild())
                {
                    TraverseLayoutBackwardLogical(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MovePreviousLogical());
        }

        /// <summary>
        /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent
        /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document
        /// </summary>
        private void PrintCurrentEntity(LayoutEnumerator layoutEnumerator, int indent)
        {
            string tabs = new String('\t', indent);

            if (layoutEnumerator.Kind == String.Empty)
            {
                Console.WriteLine($"{tabs}-> Entity type: {layoutEnumerator.Type}");
            }
            else
            {
                Console.WriteLine($"{tabs}-> Entity type & kind: {layoutEnumerator.Type}, {layoutEnumerator.Kind}");
            }

            if (layoutEnumerator.Type == LayoutEntityType.Span)
            {
                Console.WriteLine($"{tabs}   Span contents: \"{layoutEnumerator.Text}\"");
            }

            RectangleF leRect = layoutEnumerator.Rectangle;
            Console.WriteLine($"{tabs}   Rectangle dimensions {leRect.Width}x{leRect.Height}, X={leRect.X} Y={leRect.Y}");
            Console.WriteLine($"{tabs}   Page {layoutEnumerator.PageIndex}");
        }
        //ExEnd

        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void AlwaysCompressMetafiles(bool isAlwaysCompressMetafiles)
        {
            //ExStart
            //ExFor:DocSaveOptions.AlwaysCompressMetafiles
            //ExSummary:Shows how to change metafiles compression in a document while saving.
            // The document has a mathematical formula
            Document doc = new Document(MyDir + "Document.AlwaysCompressMetafiles.doc");
            
            // Large metafiles are always compressed when exporting a document in Aspose.Words, but small metafiles are not
            // compressed for performance reason. Some other document editors, such as LibreOffice, cannot read uncompressed
            // metafiles. The following option 'AlwaysCompressMetafiles' was introduced to choose appropriate behavior
            DocSaveOptions saveOptions = new DocSaveOptions();
            // False - small metafiles are not compressed for performance reason
            // True - all metafiles are compressed regardless of its size
            saveOptions.AlwaysCompressMetafiles = isAlwaysCompressMetafiles;
            
            doc.Save(ArtifactsDir + "Document.AlwaysCompressMetafiles.doc", saveOptions);
            //ExEnd
        }

        [Test]
        public void ReadMacrosFromDocument()
        {
            //ExStart
            //ExFor:Document.VbaProject
            //ExFor:VbaProject
            //ExFor:VbaModuleCollection
            //ExFor:VbaModule
            //ExFor:VbaProject.Name
            //ExFor:VbaProject.Modules
            //ExFor:VbaModule.Name
            //ExFor:VbaModule.SourceCode
            //ExSummary:Shows how to get access to VBA project information in the document.
            Document doc = new Document(MyDir + "Document.TestButton.docm");

            // A VBA project inside the document is defined as a collection of VBA modules
            VbaProject vbaProject = doc.VbaProject;
            Console.WriteLine($"Project name: {vbaProject.Name}; Modules count: {vbaProject.Modules.Count()}\n");
            
            Assert.AreEqual(vbaProject.Name, "AsposeVBAtest"); //ExSkip
            Assert.AreEqual(vbaProject.Modules.Count(), 3); //ExSkip

            VbaModuleCollection vbaModules = doc.VbaProject.Modules;
            foreach (VbaModule module in vbaModules)
            {
                Console.WriteLine($"Module name: {module.Name};\nModule code:\n{module.SourceCode}\n");
            }
            //ExEnd

            VbaModule defaultModule = vbaModules[0];
            Assert.AreEqual(defaultModule.Name, "ThisDocument");
            Assert.IsTrue(defaultModule.SourceCode.Contains("MsgBox \"First test\""));

            VbaModule createdModule = vbaModules[1];
            Assert.AreEqual(createdModule.Name, "Module1");
            Assert.IsTrue(createdModule.SourceCode.Contains("MsgBox \"Second test\""));

            VbaModule classModule = vbaModules[2];
            Assert.AreEqual(classModule.Name, "Class1");
            Assert.IsTrue(classModule.SourceCode.Contains("MsgBox \"Class test\""));
        }

#if !(NETSTANDARD2_0 || __MOBILE__)
        [Test]
        public void OpenType()
        {
            //ExStart
            //ExFor:LayoutOptions.TextShaperFactory
            //ExSummary:Shows how to support OpenType features using HarfBuzz text shaping engine.
            // Open a document
            Document doc = new Document(MyDir + "OpenType.Document.docx");

            // Please note that text shaping is only performed when exporting to PDF or XPS formats now

            // Aspose.Words is capable of using text shaper objects provided externally.
            // A text shaper represents a font and computes shaping information for a text.
            // A document typically refers to multiple fonts thus a text shaper factory is necessary.
            // When text shaper factory is set, layout starts to use OpenType features.
            // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory
            doc.LayoutOptions.TextShaperFactory = HarfBuzzTextShaperFactory.Instance;

            // Render the document to PDF format
            doc.Save(ArtifactsDir + "OpenType.Document.pdf");
            //ExEnd
        }
#endif
    }
}