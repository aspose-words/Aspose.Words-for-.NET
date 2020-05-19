// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Fonts;
using Aspose.Words.Layout;
using Aspose.Words.Lists;
using Aspose.Words.Markup;
using Aspose.Words.Properties;
using Aspose.Words.Rendering;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using Aspose.Words.Tables;
using Aspose.Words.WebExtensions;
using NUnit.Framework;
using CompareOptions = Aspose.Words.CompareOptions;
#if NET462 || NETCOREAPP2_1 || JAVA
using Aspose.Words.Shaping.HarfBuzz;
#endif
#if NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif
#if NET462 || MAC || JAVA
using Aspose.Words.Loading;
using Org.BouncyCastle.Pkcs;
using System.Security;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDocument : ApiExampleBase
    {
#if NET462 || NETCOREAPP2_1 || JAVA
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
            //ExSummary:Aspose.Words will attempt to find the license file in the embedded resources or in the assembly folders.
            License license = new License();
            license.SetLicense("Aspose.Words.lic");
            //ExEnd

            // Cleanup by removing the license
            license.SetLicense("");
            File.Delete(dstFileName);
        }

        [Test]
        public void LicenseFromStream()
        {
            // This is where the test license is on my development machine
            string testLicenseFileName = Path.Combine(LicenseDir, "Aspose.Words.lic");

            Stream myStream = File.OpenRead(testLicenseFileName);
            try
            {
                //ExStart
                //ExFor:License.SetLicense(Stream)
                //ExSummary:Shows how to initialize a license from a stream.
                License license = new License();
                license.SetLicense(myStream);
                //ExEnd
            }
            finally
            {
                myStream.Close();
            }
        }

        [Test, Category("IgnoreOnJenkins")]
        public void OpenType()
        {
            //ExStart
            //ExFor:LayoutOptions.TextShaperFactory
            //ExSummary:Shows how to support OpenType features using HarfBuzz text shaping engine.
            // Open a document
            Document doc = new Document(MyDir + "OpenType text shaping.docx");

            // Please note that text shaping is only performed when exporting to PDF or XPS formats now

            // Aspose.Words is capable of using text shaper objects provided externally
            // A text shaper represents a font and computes shaping information for a text
            // A document typically refers to multiple fonts thus a text shaper factory is necessary
            // When text shaper factory is set, layout starts to use OpenType features
            // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory
            doc.LayoutOptions.TextShaperFactory = HarfBuzzTextShaperFactory.Instance;

            // Render the document to PDF format
            doc.Save(ArtifactsDir + "Document.OpenType.pdf");
            //ExEnd
        }
#endif

#if NET462 || MAC || JAVA
        //ExStart
        //ExFor:LoadOptions.ResourceLoadingCallback
        //ExSummary:Shows how to handle external resources in Html documents during loading.
        [Test] //ExSkip
        public void LoadOptionsCallback()
        {
            // Create a new LoadOptions object and set its ResourceLoadingCallback attribute
            // as an instance of our IResourceLoadingCallback implementation 
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ResourceLoadingCallback = new HtmlLinkedResourceLoadingCallback();
            
            // When we open an Html document, external resources such as references to CSS stylesheet files and external images
            // will be handled in a custom manner by the loading callback as the document is loaded
            Document doc = new Document(MyDir + "Images.html", loadOptions);
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

                        const string newImageFilename = "Logo.jpg";
                        Console.WriteLine($"\tImage will be substituted with: {newImageFilename}");

                        Image newImage = Image.FromFile(ImageDir + newImageFilename);

                        ImageConverter converter = new ImageConverter();
                        byte[] imageBytes = (byte[])converter.ConvertTo(newImage, typeof(byte[]));
                        args.SetData(imageBytes);

                        return ResourceLoadingAction.UserProvided;
                }

                return ResourceLoadingAction.Default;
            }
        }
        //ExEnd

        [Test]
        public void CertificateHolderCreate()
        {
            //ExStart
            //ExFor:CertificateHolder.Create(Byte[], SecureString)
            //ExFor:CertificateHolder.Create(Byte[], String)
            //ExFor:CertificateHolder.Create(String, String, String)
            //ExSummary:Shows how to create CertificateHolder objects.
            // Load a PKCS #12 file into a byte array and apply its password to create the CertificateHolder
            byte[] certBytes = File.ReadAllBytes(MyDir + "morzal.pfx");
            CertificateHolder.Create(certBytes, "aw");

            // Pass a SecureString which contains the password instead of a normal string
            SecureString password = new NetworkCredential("", "aw").SecurePassword;
            CertificateHolder.Create(certBytes, password);

            // If the certificate has private keys corresponding to aliases, we can use the aliases to fetch their respective keys
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

#if NETCOREAPP2_1
        [Test]
        public void Pdf2Word()
        {
            // Check that PDF document format detects correctly
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Pdf Document.pdf");
            Assert.AreEqual(info.LoadFormat, Aspose.Words.LoadFormat.Pdf);

            // Check that PDF document opens correctly
            Document doc = new Document(MyDir + "Pdf Document.pdf");
            Assert.AreEqual(
                "Heading 1\rHeading 1.1.1.1 Heading 1.1.1.2\rHeading 1.1.1.1.1.1.1.1.1 Heading 1.1.1.1.1.1.1.1.2\u000c",
                doc.Range.Text);

            // Check that protected PDF document opens correctly
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EncryptionDetails = new PdfEncryptionDetails("Aspose", null, PdfEncryptionAlgorithm.RC4_40);

            doc.Save(ArtifactsDir + "Document.PdfDocumentEncrypted.pdf", saveOptions);

            PdfLoadOptions loadOptions = new PdfLoadOptions();
            loadOptions.Password = "Aspose";
            loadOptions.LoadFormat = Aspose.Words.LoadFormat.Pdf;

            doc = new Document(ArtifactsDir + "Document.PdfDocumentEncrypted.pdf", loadOptions);
        }
#endif

        [Test]
        public void DocumentCtor()
        {
            //ExStart
            //ExFor:Document.#ctor(Boolean)
            //ExSummary:Shows how to create a blank document.
            // Create a blank document, which will contain a section, body and paragraph by default
            Document doc = new Document();

            // Create a document object from an existing document in the local file system
            doc = new Document(MyDir + "Document.docx");

            Assert.AreEqual("Hello World!", doc.FirstSection.Body.FirstParagraph.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void ConvertToPdf()
        {
            //ExStart
            //ExFor:Document.#ctor(String)
            //ExFor:Document.Save(String)
            //ExSummary:Shows how to open a document and convert it to .PDF.
            // Open a document that exists in the local file system
            Document doc = new Document(MyDir + "Document.docx");

            // Save that document as a PDF to another location
            doc.Save(ArtifactsDir + "Document.ConvertToPdf.pdf");
            //ExEnd
        }

        [Test]
        public void OpenAndSaveToFile()
        {
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "Document.OpenAndSaveToFile.html");
        }

        [Test]
        public void OpenFromStream()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream)
            //ExSummary:Shows how to open a document from a stream.
            // Open the stream. Read only access is enough for Aspose.Words to load a document.
            using (Stream stream = File.OpenRead(MyDir + "Document.docx"))
            {
                // Load the entire document into memory and read its contents
                Document doc = new Document(stream);

                Assert.AreEqual("Hello World!", doc.GetText().Trim());
            }
            //ExEnd
        }

        [Test]
        public void OpenFromStreamWithBaseUri()
        {
            Document doc;

            //ExStart
            //ExFor:Document.#ctor(Stream,LoadOptions)
            //ExFor:LoadOptions.#ctor
            //ExFor:LoadOptions.BaseUri
            //ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
            // Open the stream
            using (Stream stream = File.OpenRead(MyDir + "Document.html"))
            {
                // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found
                // Note the Document constructor detects HTML format automatically
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.BaseUri = ImageDir;

                doc = new Document(stream, loadOptions);
            }
            //ExEnd

            // Save in the DOC format
            doc.Save(ArtifactsDir + "Document.OpenFromStreamWithBaseUri.doc");
            
            // Lets make sure the image was imported successfully into a Shape node
            // Get the first shape node in the document
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            // Verify some properties of the image
            Assert.IsTrue(shape.IsImage);
            Assert.IsNotNull(shape.ImageData.ImageBytes);
            Assert.AreEqual(32.0, ConvertUtil.PointToPixel(shape.Width), 0.01);
            Assert.AreEqual(32.0, ConvertUtil.PointToPixel(shape.Height), 0.01);
        }

        [Test]
        public void OpenDocumentFromWeb()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream)
            //ExSummary:Shows how to retrieve a document from a URL and saves it to disk in a different format.
            // This is the URL address pointing to where to find the document
            const string url = "https://omextemplates.content.office.net/support/templates/en-us/tf16402488.dotx";

            // The easiest way to load our document from the internet is make use of the 
            // System.Net.WebClient class. Create an instance of it and pass the URL
            // to download from.
            using (WebClient webClient = new WebClient())
            {
                // Download the bytes from the location referenced by the URL
                byte[] dataBytes = webClient.DownloadData(url);
                Assert.That(dataBytes, Is.Not.Empty); //ExSkip

                // Wrap the bytes representing the document in memory into a MemoryStream object
                using (MemoryStream byteStream = new MemoryStream(dataBytes))
                {
                    // Load this memory stream into a new Aspose.Words Document
                    // The file format of the passed data is inferred from the content of the bytes itself
                    // You can load any document format supported by Aspose.Words in the same way
                    Document doc = new Document(byteStream);
                    Assert.True(doc.GetText().Contains("First Name last name")); //ExSkip

                    // Convert the document to any format supported by Aspose.Words and save
                    doc.Save(ArtifactsDir + "Document.OpenDocumentFromWeb.docx");
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
            const string url = "http://www.aspose.com/";
            
            // Create a WebClient object to easily extract the HTML from the page
            WebClient client = new WebClient();
            string pageSource = client.DownloadString(url);
            client.Dispose();

            // Get the HTML as bytes for loading into a stream
            Encoding encoding = client.Encoding;
            byte[] pageBytes = encoding.GetBytes(pageSource);

            // Load the HTML into a stream
            using (MemoryStream stream = new MemoryStream(pageBytes))
            {
                // The baseUri property should be set to ensure any relative img paths are retrieved correctly
                LoadOptions options = new LoadOptions(Aspose.Words.LoadFormat.Html, "", url);

                // Load the HTML document from stream and pass the LoadOptions object
                Document doc = new Document(stream, options);

                // Save the document to the local file system while converting it to .docx
                doc.Save(ArtifactsDir + "Document.InsertHtmlFromWebPage.docx");
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
            //ExSummary:Shows how to load a document as HTML without automatic file format detection.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFormat = Aspose.Words.LoadFormat.Html;

            Document doc = new Document(MyDir + "Document.html", loadOptions);
            //ExEnd

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
        }

        [Test]
        public void LoadEncrypted()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream,LoadOptions)
            //ExFor:Document.#ctor(String,LoadOptions)
            //ExFor:LoadOptions
            //ExFor:LoadOptions.#ctor(String)
            //ExSummary:Shows how to load a Microsoft Word document encrypted with a password.
            // If we try open an encrypted document without the password, an IncorrectPasswordException will be thrown
            // We can construct a LoadOptions object with the correct encryption password
            LoadOptions options = new LoadOptions("docPassword");

            // Then, we can use that object as a parameter when opening an encrypted document
            Document doc = new Document(MyDir + "Encrypted.docx", options);
            Assert.AreEqual("Test encrypted document.", doc.GetText().Trim()); //ExSkip

            using (Stream stream = File.OpenRead(MyDir + "Encrypted.docx"))
            {
                doc = new Document(stream, options);
                Assert.AreEqual("Test encrypted document.", doc.GetText().Trim()); //ExSkip
            }
            //ExEnd

            Assert.Throws<IncorrectPasswordException>(() => doc = new Document(MyDir + "Encrypted.docx"));
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ConvertShapeToOfficeMath(bool isConvertShapeToOfficeMath)
        {
            //ExStart
            //ExFor:LoadOptions.ConvertShapeToOfficeMath
            //ExSummary:Shows how to convert shapes with EquationXML to Office Math objects.
            LoadOptions loadOptions = new LoadOptions();
            // Use 'true/false' values to convert shapes with EquationXML to Office Math objects or not
            loadOptions.ConvertShapeToOfficeMath = isConvertShapeToOfficeMath;
            
            // Specify load option to convert math shapes to office math objects on loading stage
            Document doc = new Document(MyDir + "Math shapes.docx", loadOptions);
            //ExEnd

            if (isConvertShapeToOfficeMath)
            {
                Assert.AreEqual(16, doc.GetChildNodes(NodeType.Shape, true).Count);
                Assert.AreEqual(34, doc.GetChildNodes(NodeType.OfficeMath, true).Count);
            }
            else
            {
                Assert.AreEqual(24, doc.GetChildNodes(NodeType.Shape, true).Count);
                Assert.AreEqual(0, doc.GetChildNodes(NodeType.OfficeMath, true).Count);
            }
        }

        [Test]
        public void LoadOptionsEncoding()
        {
            //ExStart
            //ExFor:LoadOptions.Encoding
            //ExSummary:Shows how to set the encoding with which to open a document.
            // Get the file format info of a file in our local file system
            FileFormatInfo fileFormatInfo = FileFormatUtil.DetectFileFormat(MyDir + "Encoded in UTF-7.txt");

            // A FileFormatInfo object can detect the encoding of the text content in a file, but in some cases it may be ambiguous
            // We know that the above file is encoded in UTF-7, but the text could be valid in others
            Assert.AreNotEqual(Encoding.UTF7, fileFormatInfo.Encoding);

            // If we open the document normally, the wrong encoding will be applied,
            // and the content of the document will not be represented correctly
            Document doc = new Document(MyDir + "Encoded in UTF-7.txt");
            Assert.AreEqual("Hello world+ACE-\r\n\r\n", doc.ToString(SaveFormat.Text));

            // In these cases we can set the Encoding attribute in a LoadOptions object
            // to override the automatically chosen encoding with the one we know to be correct
            LoadOptions loadOptions = new LoadOptions { Encoding = Encoding.UTF7 };
            doc = new Document(MyDir + "Encoded in UTF-7.txt", loadOptions);

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
            fontSettings.SetFontsFolder(FontsDir, false);
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Times New Roman", "Arvo");

            // Set that FontSettings object as a member of a newly created LoadOptions object
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;

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
            // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
            LoadOptions loadOptions = new LoadOptions();
            Assert.AreEqual(MsWordVersion.Word2019, loadOptions.MswVersion);

            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            Assert.AreEqual(12.95, doc.Styles.DefaultParagraphFormat.LineSpacing, 0.005f);

            // We can change the loading version like this, to Microsoft Word 2007
            loadOptions.MswVersion = MsWordVersion.Word2007;

            // This document is missing the default paragraph format style,
            // so when it is opened with either Microsoft Word or Aspose Words, that default style will be regenerated,
            // and will show up in the Styles collection, with values according to Microsoft Word 2007 specifications
            doc = new Document(MyDir + "Document.docx", loadOptions);
            Assert.AreEqual(13.8, doc.Styles.DefaultParagraphFormat.LineSpacing, 0.005f);
            //ExEnd
        }

        //ExStart
        //ExFor:LoadOptions.WarningCallback
        //ExSummary:Shows how to print and store warnings that occur during document loading.
        [Test] //ExSkip
        public void LoadOptionsWarningCallback()
        {
            // Create a new LoadOptions object and set its WarningCallback attribute as an instance of our IWarningCallback implementation 
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.WarningCallback = new DocumentLoadingWarningCallback();

            // Warnings that occur during loading of the document will now be printed and stored
            Document doc = new Document(MyDir + "Document.docx", loadOptions);

            List<WarningInfo> warnings = ((DocumentLoadingWarningCallback)loadOptions.WarningCallback).GetWarnings();
            Assert.AreEqual(3, warnings.Count);
            TestLoadOptionsWarningCallback(warnings); //ExSkip
        }

        /// <summary>
        /// IWarningCallback that prints warnings and their details as they arise during document loading.
        /// </summary>
        private class DocumentLoadingWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                Console.WriteLine($"Warning: {info.WarningType}");
                Console.WriteLine($"\tSource: {info.Source}");
                Console.WriteLine($"\tDescription: {info.Description}");
                mWarnings.Add(info);
            }

            public List<WarningInfo> GetWarnings()
            {
                return mWarnings;
            }

            private readonly List<WarningInfo> mWarnings = new List<WarningInfo>();
        }
        //ExEnd

        private static void TestLoadOptionsWarningCallback(List<WarningInfo> warnings)
        {
            Assert.AreEqual(WarningType.UnexpectedContent, warnings[0].WarningType);
            Assert.AreEqual(WarningSource.Docx, warnings[0].Source);
            Assert.AreEqual("3F01", warnings[0].Description);

            Assert.AreEqual(WarningType.MinorFormattingLoss, warnings[1].WarningType);
            Assert.AreEqual(WarningSource.Docx, warnings[1].Source);
            Assert.AreEqual("Import of element 'shapedefaults' is not supported in Docx format by Aspose.Words.", warnings[1].Description); 

            Assert.AreEqual(WarningType.MinorFormattingLoss, warnings[2].WarningType); 
            Assert.AreEqual(WarningSource.Docx, warnings[2].Source);
            Assert.AreEqual("Import of element 'extraClrSchemeLst' is not supported in Docx format by Aspose.Words.", warnings[2].Description); 
        }

        [Test]
        public void ConvertToHtml()
        {
            //ExStart
            //ExFor:Document.Save(String,SaveFormat)
            //ExFor:SaveFormat
            //ExSummary:Shows how to convert from DOCX to HTML format.
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "Document.ConvertToHtml.html", SaveFormat.Html);
            //ExEnd
        }

        [Test]
        public void ConvertToMhtml()
        {
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "Document.ConvertToMhtml.mht");
        }

        [Test]
        public void ConvertToTxt()
        {
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "Document.ConvertToTxt.txt");
            
        }

        [Test]
        public void SaveToStream()
        {
            //ExStart
            //ExFor:Document.Save(Stream,SaveFormat)
            //ExSummary:Shows how to save a document to a stream.
            Document doc = new Document(MyDir + "Document.docx");

            using (MemoryStream dstStream = new MemoryStream())
            {
                doc.Save(dstStream, SaveFormat.Docx);

                // Rewind the stream position back to zero so it is ready for next reader
                dstStream.Position = 0;
                Assert.AreEqual("Hello World!", new Document(dstStream).GetText().Trim()); //ExSkip
            }
            //ExEnd
        }

        [Test]
        public void Doc2EpubSave()
        {
            // Open an existing document from disk
            Document doc = new Document(MyDir + "Rendering.docx");

            // Save the document in EPUB format
            doc.Save(ArtifactsDir + "Document.Doc2EpubSave.epub");
        }

        [Test]
        public void Doc2EpubSaveOptions()
        {
            //ExStart
            //ExFor:DocumentSplitCriteria
            //ExFor:HtmlSaveOptions
            //ExFor:HtmlSaveOptions.#ctor
            //ExFor:HtmlSaveOptions.Encoding
            //ExFor:HtmlSaveOptions.DocumentSplitCriteria
            //ExFor:HtmlSaveOptions.ExportDocumentProperties
            //ExFor:HtmlSaveOptions.SaveFormat
            //ExFor:SaveOptions
            //ExFor:SaveOptions.SaveFormat
            //ExSummary:Shows how to convert a document to EPUB with save options specified.
            // Open an existing document from disk
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
            // how the output document is saved
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Specify the desired encoding
            saveOptions.Encoding = Encoding.UTF8;

            // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
            // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
            // HTML files greater than a certain size e.g 300kb
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

            // Specify that we want to export document properties
            saveOptions.ExportDocumentProperties = true;

            // Specify that we want to save in EPUB format
            saveOptions.SaveFormat = SaveFormat.Epub;

            // Export the document as an EPUB file
            doc.Save(ArtifactsDir + "Document.Doc2EpubSaveOptions.epub", saveOptions);
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
            //ExFor:PdfSaveOptions.DownsampleOptions
            //ExSummary:Shows how to change the resolution of images in output pdf documents.
            // Open a document that contains images 
            Document doc = new Document(MyDir + "Rendering.docx");

            doc.Save(ArtifactsDir + "Document.DownsampleOptions.Default.pdf");

            // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
            PdfSaveOptions options = new PdfSaveOptions();

            // This conversion will downsample images by default
            Assert.True(options.DownsampleOptions.DownsampleImages);
            Assert.AreEqual(220, options.DownsampleOptions.Resolution);

            // We can set the output resolution to a different value
            // The first two images in the input document will be affected by this
            options.DownsampleOptions.Resolution = 36;

            // We can set a minimum threshold for downsampling 
            // This value will prevent some images in the input document from being downsampled
            options.DownsampleOptions.ResolutionThreshold = 128;

            doc.Save(ArtifactsDir + "Document.DownsampleOptions.LowerThreshold.pdf", options);
            //ExEnd
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void SaveHtmlPrettyFormat(bool isPrettyFormat)
        {
            //ExStart
            //ExFor:SaveOptions.PrettyFormat
            //ExSummary:Shows how to pass an option to export HTML tags in a well spaced, human readable format.
            Document doc = new Document(MyDir + "Document.docx");

            // Enabling the PrettyFormat setting will export HTML in an indented format that is easy to read
            // If this is setting is false (by default) then the HTML tags will be exported in condensed form with no indentation
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.PrettyFormat = isPrettyFormat;

            doc.Save(ArtifactsDir + "Document.SaveHtmlPrettyFormat.html", htmlOptions);
            //ExEnd

            string html = File.ReadAllText(ArtifactsDir + "Document.SaveHtmlPrettyFormat.html");

            // Enabling HtmlSaveOptions.PrettyFormat places tabs and newlines in places where it would improve the readability of html source
            Assert.True(isPrettyFormat
                ? html.StartsWith(
                    "<html>\r\n\t<head>\r\n\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n\t\t")
                : html.StartsWith(
                    "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />"));
        }

        [Test]
        public void SaveHtmlWithOptions()
        {
            //ExStart
            //ExFor:HtmlSaveOptions
            //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
            //ExFor:HtmlSaveOptions.ImagesFolder
            //ExSummary:Shows how to set save options before saving a document to HTML.
            Document doc = new Document(MyDir + "Rendering.docx");

            // This is the directory we want the exported images to be saved to
            string imagesDir = Path.Combine(ArtifactsDir, "SaveHtmlWithOptions");

            // The folder specified needs to exist and should be empty
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportTextInputFormFieldAsText = true;
            options.ImagesFolder = imagesDir;

            doc.Save(ArtifactsDir + "Document.SaveHtmlWithOptions.html", options);
            //ExEnd

            // Verify the images were saved to the correct location
            Assert.IsTrue(File.Exists(ArtifactsDir + "Document.SaveHtmlWithOptions.html"));

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
            Document doc = new Document(MyDir + "Rendering.docx");

            // Set the option to export font resources and create and pass the object which implements the handler methods
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
            options.ExportFontResources = true;
            options.FontSavingCallback = new HandleFontSaving();
            
            // The fonts from the input document will now be exported as .ttf files and saved alongside the output document
            doc.Save(ArtifactsDir + "Document.SaveHtmlExportFonts.html", options);
            Assert.AreEqual(10, Array.FindAll(Directory.GetFiles(ArtifactsDir), s => s.EndsWith(".ttf")).Length); //ExSkip
        }

        /// <summary>
        /// Prints information about fonts and saves them alongside their output .html.
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
                args.FontFileName = args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last();

                args.FontStream = 
                    new FileStream(ArtifactsDir + args.OriginalFileName.Split(Path.DirectorySeparatorChar).Last(), FileMode.Create);
                Assert.False(args.KeepFontStreamOpen);

                // We can access the source document from here also
                Assert.True(args.Document.OriginalFileName.EndsWith("Rendering.docx"));
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
        //ExSummary:Shows how to implement custom logic over node insertion in the document by changing the font of inserted HTML content.
        [Test] //ExSkip
        public void FontChangeViaCallback()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set up and pass the object which implements the handler methods
            doc.NodeChangingCallback = new HandleNodeChangingFontChanger();

            // Insert sample HTML content
            builder.InsertHtml("<p>Hello World</p>");

            doc.Save(ArtifactsDir + "Document.FontChangeViaCallback.docx");
            doc = new Document(ArtifactsDir + "Document.FontChangeViaCallback.docx"); //ExSkip
            Run run = (Run)doc.GetChild(NodeType.Run, 0, true); //ExSkip
            Assert.AreEqual(24.0, run.Font.Size); //ExSkip
            Assert.AreEqual("Arial", run.Font.Name); //ExSkip
        }

        public class HandleNodeChangingFontChanger : INodeChangingCallback
        {
            // Implement the NodeInserted handler to set default font settings for every Run node inserted into the Document
            void INodeChangingCallback.NodeInserted(NodeChangingArgs args)
            {
                // Change the font of inserted text contained in the Run nodes
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
            // The document that the content will be appended to
            Document dstDoc = new Document();
            dstDoc.FirstSection.Body.AppendParagraph("Destination document text. ");

            // The document to append
            Document srcDoc = new Document();
            srcDoc.FirstSection.Body.AppendParagraph("Source document text. ");

            // Append the source document to the destination document
            // Pass format mode to retain the original formatting of the source document when importing it
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            Assert.AreEqual(2, dstDoc.Sections.Count); //ExSkip

            // Save the document
            dstDoc.Save(ArtifactsDir + "Document.AppendDocument.docx");
            //ExEnd

            string outDocText = new Document(ArtifactsDir + "Document.AppendDocument.docx").GetText();

            Assert.True(outDocText.StartsWith(dstDoc.GetText()));
            Assert.True(outDocText.EndsWith(srcDoc.GetText()));
        }

        [Test]
        // Using this file path keeps the example making sense when compared with automation so we expect
        // the file not to be found
        public void AppendDocumentFromAutomation()
        {
            // The document that the other documents will be appended to
            Document doc = new Document();
            
            // We should call this method to clear this document of any existing content
            doc.RemoveAllChildren();

            const int recordCount = 5;
            for (int i = 1; i <= recordCount; i++)
            {
                Document srcDoc = new Document();

                // Open the document to join.
                Assert.That(() => srcDoc == new Document("C:\\DetailsList.doc"),
                    Throws.TypeOf<FileNotFoundException>());

                // Append the source document at the end of the destination document
                doc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

                // In automation you were required to insert a new section break at this point, however in Aspose.Words we 
                // don't need to do anything here as the appended document is imported as separate sections already

                // If this is the second document or above being appended then unlink all headers footers in this section 
                // from the headers and footers of the previous section
                if (i > 1)
                    Assert.That(() => doc.Sections[i].HeadersFooters.LinkToPrevious(false),
                        Throws.TypeOf<NullReferenceException>());
            }
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
            //ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
            // Load the document which contains signature
            Document doc = new Document(MyDir + "Digitally signed.docx");

            foreach (DigitalSignature signature in doc.DigitalSignatures)
            {
                Console.WriteLine("*** Signature Found ***");
                Console.WriteLine("Is valid: " + signature.IsValid);
                // This property is available in MS Word documents only
                Console.WriteLine("Reason for signing: " + signature.Comments); 
                Console.WriteLine("Signature type: " + signature.SignatureType);
                Console.WriteLine("Time of signing: " + signature.SignTime);
                Console.WriteLine("Subject name: " + signature.CertificateHolder.Certificate.SubjectName);
                Console.WriteLine("Issuer name: " + signature.CertificateHolder.Certificate.IssuerName.Name);
                Console.WriteLine();
            }
            //ExEnd

            Assert.AreEqual(1, doc.DigitalSignatures.Count);

            DigitalSignature digitalSig = doc.DigitalSignatures[0];

            Assert.True(digitalSig.IsValid);
            Assert.AreEqual("Test Sign", digitalSig.Comments);
            Assert.AreEqual("XmlDsig", digitalSig.SignatureType.ToString());
            Assert.True(digitalSig.CertificateHolder.Certificate.Subject.Contains("Aspose Pty Ltd"));
            Assert.True(digitalSig.CertificateHolder.Certificate.IssuerName.Name != null &&
                        digitalSig.CertificateHolder.Certificate.IssuerName.Name.Contains("VeriSign"));
        }

        [Test]
        public void DigitalSignature()
        {
            //ExStart
            //ExFor:DigitalSignature.CertificateHolder
            //ExFor:DigitalSignature.IssuerName
            //ExFor:DigitalSignature.SubjectName
            //ExFor:DigitalSignatureCollection
            //ExFor:DigitalSignatureCollection.IsValid
            //ExFor:DigitalSignatureCollection.Count
            //ExFor:DigitalSignatureCollection.Item(Int32)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder)
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder)
            //ExFor:DigitalSignatureType
            //ExFor:Document.DigitalSignatures
            //ExSummary:Shows how to sign documents with X.509 certificates.
            // Verify that a document isn't signed
            Assert.False(FileFormatUtil.DetectFileFormat(MyDir + "Document.docx").HasDigitalSignature);

            // Create a CertificateHolder object from a PKCS #12 file, which we will use to sign the document
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);

            // There are 2 ways of saving a signed copy of a document to the local file system
            // 1: Designate unsigned input and signed output files by filename and sign with the passed CertificateHolder 
            DigitalSignatureUtil.Sign(MyDir + "Document.docx", ArtifactsDir + "Document.DigitalSignature.docx", 
                certificateHolder, new SignOptions() { SignTime = DateTime.Now } );

            Assert.True(FileFormatUtil.DetectFileFormat(ArtifactsDir + "Document.DigitalSignature.docx").HasDigitalSignature);

            // 2: Create a stream for the input file and one for the output and create a file, signed with the CertificateHolder, at the file system location determine
            using (FileStream inDoc = new FileStream(MyDir + "Document.docx", FileMode.Open))
            {
                using (FileStream outDoc = new FileStream(ArtifactsDir + "Document.DigitalSignature.docx", FileMode.Create))
                {
                    DigitalSignatureUtil.Sign(inDoc, outDoc, certificateHolder);
                }
            }

            Assert.True(FileFormatUtil.DetectFileFormat(ArtifactsDir + "Document.DigitalSignature.docx").HasDigitalSignature);

            // Open the signed document and get its digital signature collection
            Document signedDoc = new Document(ArtifactsDir + "Document.DigitalSignature.docx");
            DigitalSignatureCollection digitalSignatureCollection = signedDoc.DigitalSignatures;

            // Verify that all of the document's digital signatures are valid and check their details
            Assert.True(digitalSignatureCollection.IsValid);
            Assert.AreEqual(1, digitalSignatureCollection.Count);
            Assert.AreEqual(DigitalSignatureType.XmlDsig, digitalSignatureCollection[0].SignatureType);
            Assert.AreEqual("CN=Morzal.Me", signedDoc.DigitalSignatures[0].IssuerName);
            Assert.AreEqual("CN=Morzal.Me", signedDoc.DigitalSignatures[0].SubjectName);
            //ExEnd
        }

        [Test]
        public void AppendAllDocumentsInFolder()
        {
            string path = ArtifactsDir + "Document.AppendAllDocumentsInFolder.doc";

            // Delete the file that was created by the previous run as I don't want to append it again
            if (File.Exists(path))
                File.Delete(path);

            //ExStart
            //ExFor:Document.AppendDocument(Document, ImportFormatMode)
            //ExSummary:Shows how to use the AppendDocument method to combine all the documents in a folder to the end of a template document.
            // Lets start with a simple template and append all the documents in a folder to this document
            Document baseDoc = new Document();

            // Add some content to the template
            DocumentBuilder builder = new DocumentBuilder(baseDoc);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Template Document");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Some content here");

            // Gather the files which will be appended to our template document
            // In this case we add the optional parameter to include the search only for files with the ".doc" extension
            ArrayList files = new ArrayList(Directory.GetFiles(MyDir, "*.doc")
                .Where(file => file.EndsWith(".doc", StringComparison.CurrentCultureIgnoreCase)).ToArray());
            Assert.AreEqual(7, files.Count); //ExSkip

            // The list of files may come in any order, let's sort the files by name so the documents are enumerated alphabetically
            files.Sort();
            Assert.AreEqual(5, baseDoc.Styles.Count); //ExSkip
            Assert.AreEqual(1, baseDoc.Sections.Count); //ExSkip

            // Iterate through every file in the directory and append each one to the end of the template document
            foreach (string fileName in files)
            {
                // We have some encrypted test documents in our directory, Aspose.Words can open encrypted documents 
                // but only with the correct password. Let's just skip them here for simplicity
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(fileName);
                if (info.IsEncrypted)
                    continue;

                Document subDoc = new Document(fileName);
                baseDoc.AppendDocument(subDoc, ImportFormatMode.UseDestinationStyles);
            }

            // Save the combined document to disk
            baseDoc.Save(path);
            //ExEnd

            Assert.AreEqual(7, baseDoc.Styles.Count);
            Assert.AreEqual(8, baseDoc.Sections.Count);
        }

        [Test]
        public void JoinRunsWithSameFormatting()
        {
            //ExStart
            //ExFor:Document.JoinRunsWithSameFormatting
            //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
            // Open a document which contains adjacent runs of text with identical formatting
            // This can, for example, occur if we edit one paragraph many times
            Document doc = new Document(MyDir + "Rendering.docx");

            // Get the number of runs our document contains
            Assert.AreEqual(317, doc.GetChildNodes(NodeType.Run, true).Count);

            // We can merge all nearby runs with the same formatting to reduce that number by calling JoinRunsWithSameFormatting()
            // This method will also notify us of the number of run joins that took place
            Assert.AreEqual(121, doc.JoinRunsWithSameFormatting());

            // Get the number of runs after joining, which, together with the number of joins should add up to the original number of runs
            Assert.AreEqual(196, doc.GetChildNodes(NodeType.Run, true).Count);
            //ExEnd
        }

        [Test]
        public void DefaultTabStop()
        {
            //ExStart
            //ExFor:Document.DefaultTabStop
            //ExFor:ControlChar.Tab
            //ExFor:ControlChar.TabChar
            //ExSummary:Shows how to change default tab positions for the document and inserts text with some tab characters.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set default tab stop to 72 points (1 inch)
            builder.Document.DefaultTabStop = 72;

            builder.Writeln("Hello" + ControlChar.Tab + "World!");
            builder.Writeln("Hello" + ControlChar.TabChar + "World!");
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            Assert.AreEqual(72, doc.DefaultTabStop);
        }

        [Test]
        public void CloneDocument()
        {
            //ExStart
            //ExFor:Document.Clone
            //ExSummary:Shows how to deep clone a document.
            Document doc = new Document(MyDir + "Document.docx");
            Document clone = doc.Clone();

            Assert.AreNotEqual(doc, clone);
            //ExEnd
        }

        [Test]
        public void ChangeFieldUpdateCultureSource()
        {
            //ExStart
            //ExFor:Document.FieldOptions
            //ExFor:FieldOptions
            //ExFor:FieldOptions.FieldUpdateCultureSource
            //ExFor:FieldUpdateCultureSource
            //ExSummary:Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert two merge fields with German locale
            builder.Font.LocaleId = 1031;
            builder.InsertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
            builder.Write(" - ");
            builder.InsertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

            // Store the current culture in a variable and explicitly set it to US English
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            // Execute a mail merge for the first MERGEFIELD using the current culture (US English) for date formatting
            doc.MailMerge.Execute(new[] { "Date1" }, new object[] { new DateTime(2020, 1, 01) });

            // Execute a mail merge for the second MERGEFIELD using the field's culture (German) for date formatting
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.MailMerge.Execute(new[] { "Date2" }, new object[] { new DateTime(2020, 1, 01) });

            // The first MERGEFIELD has received a date formatted in English, while the second one is in German
            Assert.AreEqual("Wednesday, 1 January 2020 - Mittwoch, 1 Januar 2020", doc.Range.Text.Trim());

            // Restore the original culture
            Thread.CurrentThread.CurrentCulture = currentCulture;
            //ExEnd
        }

        [Test]
        public void DocumentGetTextToString()
        {
            //ExStart
            //ExFor:CompositeNode.GetText
            //ExFor:Node.ToString(SaveFormat)
            //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
            Document doc = new Document();

            // Enter a field into the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Field");

            // GetText will retrieve all field codes and special characters
            Assert.AreEqual("\u0013MERGEFIELD Field\u0014«Field»\u0015\u000c", doc.GetText());

            // ToString will give us the plaintext version of the document in the save format we put into the parameter
            Assert.AreEqual("«Field»\r\n", doc.ToString(SaveFormat.Text));
            //ExEnd
        }

        [Test]
        public void DocumentByteArray()
        {
            // Load the document
            Document doc = new Document(MyDir + "Document.docx");

            // Create a new memory stream
            MemoryStream streamOut = new MemoryStream();
            // Save the document to stream
            doc.Save(streamOut, SaveFormat.Docx);

            // Convert the document to byte form
            byte[] docBytes = streamOut.ToArray();

            // We can load the bytes back into a document object
            MemoryStream streamIn = new MemoryStream(docBytes);

            // Load the stream into a new document object
            Document loadDoc = new Document(streamIn);
            Assert.AreEqual(doc.GetText(), loadDoc.GetText());
        }

        [Test]
        public void Protect()
        {
            //ExStart
            //ExFor:Document.Protect(ProtectionType,String)
            //ExFor:Document.ProtectionType
            //ExFor:Document.Unprotect
            //ExFor:Document.Unprotect(String)
            //ExSummary:Shows how to protect a document.
            // Create a new document and protect it with a password
            Document doc = new Document();
            doc.Protect(ProtectionType.ReadOnly, "password");
            Assert.AreEqual(ProtectionType.ReadOnly, doc.ProtectionType);

            // If we open this document with Microsoft Word and wish to edit it, 
            // we will first need to stop the protection, which can only be done with the password
            doc.Save(ArtifactsDir + "Document.Protect.docx");

            // Note that the protection only applies to Microsoft Word users opening out document
            // The document can still be opened and edited programmatically without a password, despite its protection status
            // Encryption offers a more robust option for protecting document content
            Document protectedDoc = new Document(ArtifactsDir + "Document.Protect.docx");
            Assert.AreEqual(ProtectionType.ReadOnly, protectedDoc.ProtectionType);

            DocumentBuilder builder = new DocumentBuilder(protectedDoc);
            builder.Writeln("Text added to a protected document.");
            Assert.AreEqual("Text added to a protected document.", protectedDoc.Range.Text.Trim()); //ExSkip

            // Documents can have protection removed either with no password, or with the correct password
            doc.Unprotect();
            Assert.AreEqual(ProtectionType.NoProtection, doc.ProtectionType);

            doc.Protect(ProtectionType.ReadOnly, "newPassword");
            doc.Unprotect("wrongPassword"); //ExSkip
            Assert.AreEqual(ProtectionType.ReadOnly, doc.ProtectionType); //ExSkip
            doc.Unprotect("newPassword");
            Assert.AreEqual(ProtectionType.NoProtection, doc.ProtectionType);
            //ExEnd
        }

        [Test]
        public void DocumentEnsureMinimum()
        {
            //ExStart
            //ExFor:Document.EnsureMinimum
            //ExSummary:Shows how to ensure the Document is valid (has the minimum nodes required to be valid).
            Document doc = new Document();

            // Every blank document that we create will contain
            // the minimal set nodes requited for editing; a Section, Body and Paragraph
            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Any, true).Count);

            // We can remove every node from the document with RemoveAllChildren()
            doc.RemoveAllChildren();
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Any, true).Count);

            // EnsureMinimum() can ensure that the document has at least those three nodes
            doc.EnsureMinimum();
            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Any, true).Count);
            //ExEnd

            NodeCollection nodes = doc.GetChildNodes(NodeType.Any, true);

            Assert.AreEqual(NodeType.Section, nodes[0].NodeType);
            Assert.AreEqual(NodeType.Body, nodes[1].NodeType);
            Assert.AreEqual(NodeType.Paragraph, nodes[2].NodeType);

            Assert.True(nodes[1].ParentNode == nodes[0]);
            Assert.True(nodes[2].ParentNode == nodes[1]);
        }

        [Test]
        public void RemoveMacrosFromDocument()
        {
            //ExStart
            //ExFor:Document.RemoveMacros
            //ExSummary:Shows how to remove all macros from a document.
            // Open a document that contains a VBA project and macros
            Document doc = new Document(MyDir + "Macro.docm");

            Assert.IsTrue(doc.HasMacros);
            Assert.AreEqual("Project", doc.VbaProject.Name); //ExSkip

            // We can strip the document of this content by calling this method
            doc.RemoveMacros();

            Assert.IsFalse(doc.HasMacros);
            Assert.Null(doc.VbaProject); //ExSkip
            //ExEnd
        }

        [Test]
        public void UpdateTableLayout()
        {
            //ExStart
            //ExFor:Document.UpdateTableLayout
            //ExSummary:Shows how to update the layout of tables in a document.
            // Create a new document and insert a table
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.InsertCell();
            builder.Write("Cell 3");

            // Create a SaveOptions object to prepare this document to be saved to .txt
            TxtSaveOptions options = new TxtSaveOptions();
            options.PreserveTableLayout = true;
        
            // Previewing the appearance of the document in .txt form shows that the table will not be represented accurately
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true); //ExSkip
            Assert.AreEqual(0.0d, table.FirstRow.Cells[0].CellFormat.Width); //ExSkip
            Assert.AreEqual("CCC\r\neee\r\nlll\r\nlll\r\n   \r\n123\r\n\r\n", doc.ToString(options));

            // We can call UpdateTableLayout() to fix some of these issues
            doc.UpdateTableLayout();

            Assert.AreEqual(155.65d, table.FirstRow.Cells[0].CellFormat.Width); //ExSkip
            Assert.AreEqual("Cell 1             Cell 2             Cell 3\r\n\r\n", doc.ToString(options));
            //ExEnd
        }

        [Test]
        public void GetPageCount()
        {
            //ExStart
            //ExFor:Document.PageCount
            //ExSummary:Shows how to invoke page layout and retrieve the number of pages in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text spanning 3 pages
            builder.Write("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Page 3");

            // Get the page count
            Assert.AreEqual(3, doc.PageCount);

            // Getting the PageCount property invoked the document's page layout to calculate the value
            // This operation will not need to be re-done when rendering the document to a save format like .pdf,
            // which can save time with larger documents
            doc.Save(ArtifactsDir + "Document.GetPageCount.pdf");
            //ExEnd
        }

        [Test]
        public void GetUpdatedPageProperties()
        {
            //ExStart
            //ExFor:Document.UpdateWordCount()
            //ExFor:Document.UpdateWordCount(Boolean)
            //ExFor:BuiltInDocumentProperties.Characters
            //ExFor:BuiltInDocumentProperties.Words
            //ExFor:BuiltInDocumentProperties.Paragraphs
            //ExFor:BuiltInDocumentProperties.Lines
            //ExSummary:Shows how to update all list labels in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Add a paragraph of text to the document
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.Write("Ut enim ad minim veniam, " +
                            "quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

            // Document metrics are not tracked in code in real time
            Assert.AreEqual(0, doc.BuiltInDocumentProperties.Characters);
            Assert.AreEqual(0, doc.BuiltInDocumentProperties.Words);
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.Paragraphs);
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.Lines);

            // We will need to call this method to update them
            doc.UpdateWordCount();

            // Check the values of the properties
            Assert.AreEqual(196, doc.BuiltInDocumentProperties.Characters);
            Assert.AreEqual(36, doc.BuiltInDocumentProperties.Words);
            Assert.AreEqual(2, doc.BuiltInDocumentProperties.Paragraphs);
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.Lines);

            // To also get the line count as it would appear in Microsoft Word,
            // we will need to pass "true" to UpdateWordCount()
            doc.UpdateWordCount(true);
            Assert.AreEqual(4, doc.BuiltInDocumentProperties.Lines);
            //ExEnd
        }

        [Test]
        public void TableStyleToDirectFormatting()
        {
            //ExStart
            //ExFor:Document.ExpandTableStylesToDirectFormatting
            //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
            Document doc = new Document(MyDir + "Tables.docx");
            
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            // First print the color of the cell shading. This should be empty as the current shading
            // is stored in the table style
            double cellShadingBefore = table.FirstRow.RowFormat.Height;
            Console.WriteLine("Cell shading before style expansion: " + cellShadingBefore);

            // Expand table style formatting to direct formatting
            doc.ExpandTableStylesToDirectFormatting();

            // Now print the cell shading after expanding table styles. A blue background pattern color
            // should have been applied from the table style
            double cellShadingAfter = table.FirstRow.RowFormat.Height;
            Console.WriteLine("Cell shading after style expansion: " + cellShadingAfter);

            doc.Save(ArtifactsDir + "Document.TableStyleToDirectFormatting.docx");
            //ExEnd

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
            Document doc = new Document(MyDir + "Document.docx");

            // This property will return the full path and file name where the document was loaded from
            Assert.AreEqual(MyDir + "Document.docx", doc.OriginalFileName);

            // This is the original LoadFormat of the document
            Assert.AreEqual(Aspose.Words.LoadFormat.Docx, doc.OriginalLoadFormat);
            //ExEnd
        }

        [Test]
        [Description("WORDSNET-16099")]
        public void FootnoteColumns()
        {
            //ExStart
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.Columns
            //ExSummary:Shows how to set the number of columns with which the footnotes area is formatted.
            Document doc = new Document(MyDir + "Footnotes and endnotes.docx");
            Assert.AreEqual(0, doc.FootnoteOptions.Columns); //ExSkip

            // Let's change number of columns for footnotes on page. If columns value is 0 than footnotes area
            // is formatted with a number of columns based on the number of columns on the displayed page
            doc.FootnoteOptions.Columns = 2;
            doc.Save(ArtifactsDir + "Document.FootnoteColumns.docx");
            //ExEnd

            // Assert that number of columns gets correct
            doc = new Document(ArtifactsDir + "Document.FootnoteColumns.docx");

            Assert.AreEqual(2, doc.FirstSection.PageSetup.FootnoteOptions.Columns);
        }

        [Test]
        public void Footnotes()
        {
            //ExStart
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.NumberStyle
            //ExFor:FootnoteOptions.Position
            //ExFor:FootnoteOptions.RestartRule
            //ExFor:FootnoteOptions.StartNumber
            //ExFor:FootnoteNumberingRule
            //ExFor:FootnotePosition
            //ExSummary:Shows how to insert footnotes and edit their appearance.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert 3 paragraphs with a footnote at the end of each one
            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 2");
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote 3", "Custom reference mark");

            // Edit the numbering and positioning of footnotes 
            doc.FootnoteOptions.Position = FootnotePosition.BeneathText;
            doc.FootnoteOptions.NumberStyle = NumberStyle.UppercaseRoman;
            doc.FootnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;
            doc.FootnoteOptions.StartNumber = 1;

            doc.Save(ArtifactsDir + "Document.Footnotes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.Footnotes.docx");

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, 
                "Footnote 1", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty, 
                "Footnote 2", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
            TestUtil.VerifyFootnote(FootnoteType.Footnote, false, "Custom reference mark", 
                "Custom reference mark Footnote 3", (Footnote)doc.GetChild(NodeType.Footnote, 2, true));
        }

        [Test]
        public void Endnotes()
        {
            //ExStart
            //ExFor:Document.EndnoteOptions
            //ExFor:EndnoteOptions
            //ExFor:EndnoteOptions.NumberStyle
            //ExFor:EndnoteOptions.Position
            //ExFor:EndnoteOptions.RestartRule
            //ExFor:EndnoteOptions.StartNumber
            //ExFor:EndnotePosition
            //ExSummary:Shows how to insert endnotes and edit their appearance.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert 3 paragraphs with an endnote at the end of each one
            builder.Write("Text 1. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 1");
            builder.Write("Text 2. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Text 3. ");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote 3", "Custom reference mark");

            Assert.AreEqual(1, doc.EndnoteOptions.StartNumber); //ExSkip
            Assert.AreEqual(EndnotePosition.EndOfDocument, doc.EndnoteOptions.Position); //ExSkip
            Assert.AreEqual(NumberStyle.LowercaseRoman, doc.EndnoteOptions.NumberStyle); //ExSkip
            Assert.AreEqual(FootnoteNumberingRule.Default, doc.EndnoteOptions.RestartRule); //ExSkip
            
            // Edit the numbering and positioning of endnotes
            doc.EndnoteOptions.Position = EndnotePosition.EndOfDocument;
            doc.EndnoteOptions.NumberStyle = NumberStyle.UppercaseRoman;
            doc.EndnoteOptions.RestartRule = FootnoteNumberingRule.Continuous;
            doc.EndnoteOptions.StartNumber = 1;

            doc.Save(ArtifactsDir + "Document.Endnotes.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.Endnotes.docx");

            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 1", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "Endnote 2", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, false, "Custom reference mark",
                "Custom reference mark Endnote 3", (Footnote)doc.GetChild(NodeType.Footnote, 2, true));
        }

        [Test]
        public void Compare()
        {
            //ExStart
            //ExFor:Document.Compare(Document, String, DateTime)
            //ExFor:RevisionCollection.AcceptAll
            //ExSummary:Shows how to apply the compare method to two documents and then use the results. 
            Document doc1 = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc1);
            builder.Writeln("This is the original document.");

            Document doc2 = new Document();
            builder = new DocumentBuilder(doc2);
            builder.Writeln("This is the edited document.");

            // If either document has a revision, an exception will be thrown
            if (doc1.Revisions.Count == 0 && doc2.Revisions.Count == 0)
                doc1.Compare(doc2, "authorName", DateTime.Now);

            // If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed
            Assert.AreEqual(2, doc1.Revisions.Count); //ExSkip
            foreach (Revision r in doc1.Revisions)
            {
                Console.WriteLine($"Revision type: {r.RevisionType}, on a node of type \"{r.ParentNode.NodeType}\"");
                Console.WriteLine($"\tChanged text: \"{r.ParentNode.GetText()}\"");
            }

            // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2
            doc1.Revisions.AcceptAll();

            // doc1, when saved, now resembles doc2
            doc1.Save(ArtifactsDir + "Document.Compare.docx");
            //ExEnd

            doc1 = new Document(ArtifactsDir + "Document.Compare.docx");
            Assert.AreEqual(0, doc1.Revisions.Count);
            Assert.AreEqual(doc2.GetText().Trim(), doc1.GetText().Trim());
        }

        [Test]
        public void CompareDocumentWithRevisions()
        {
            Document doc1 = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc1);
            builder.Writeln("Hello world! This text is not a revision.");

            Document docWithRevision = new Document();
            builder = new DocumentBuilder(docWithRevision);

            docWithRevision.StartTrackRevisions("John Doe");
            builder.Writeln("This is a revision.");

            Assert.That(() => docWithRevision.Compare(doc1, "John Doe", DateTime.Now),
                Throws.TypeOf<InvalidOperationException>());
        }

        [Test]
        public void CompareOptions()
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
            //ExSummary:Shows how to specify which document shall be used as a target during comparison.
            // Create our original document
            Document docOriginal = new Document();
            DocumentBuilder builder = new DocumentBuilder(docOriginal);

            // Insert paragraph text with an endnote
            builder.Writeln("Hello world! This is the first paragraph.");
            builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

            // Insert a table
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Original cell 1 text");
            builder.InsertCell();
            builder.Write("Original cell 2 text");
            builder.EndTable();

            // Insert a textbox
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 20);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Original textbox contents");

            // Insert a DATE field
            builder.MoveTo(docOriginal.FirstSection.Body.AppendParagraph(""));
            builder.InsertField(" DATE ");

            // Insert a comment
            Comment newComment = new Comment(docOriginal, "John Doe", "J.D.", DateTime.Now);
            newComment.SetText("Original comment.");
            builder.CurrentParagraph.AppendChild(newComment);

            // Insert a header
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Original header contents.");

            // Create a clone of our document, which we will edit and later compare to the original
            Document docEdited = (Document)docOriginal.Clone(true);
            Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;

            // Change the formatting of the first paragraph, change casing of original characters and add text
            firstParagraph.Runs[0].Text = "hello world! this is the first paragraph, after editing.";
            firstParagraph.ParagraphFormat.Style = docEdited.Styles[StyleIdentifier.Heading1];
            
            // Edit the footnote
            Footnote footnote = (Footnote)docEdited.GetChild(NodeType.Footnote, 0, true);
            footnote.FirstParagraph.Runs[1].Text = "Edited endnote text.";

            // Edit the table
            Table table = (Table)docEdited.GetChild(NodeType.Table, 0, true);
            table.FirstRow.Cells[1].FirstParagraph.Runs[0].Text = "Edited Cell 2 contents";

            // Edit the textbox
            textBox = (Shape)docEdited.GetChild(NodeType.Shape, 0, true);
            textBox.FirstParagraph.Runs[0].Text = "Edited textbox contents";

            // Edit the DATE field
            FieldDate fieldDate = (FieldDate)docEdited.Range.Fields[0];
            fieldDate.UseLunarCalendar = true;

            // Edit the comment
            Comment comment = (Comment)docEdited.GetChild(NodeType.Comment, 0, true);
            comment.FirstParagraph.Runs[0].Text = "Edited comment.";

            // Edit the header
            docEdited.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].FirstParagraph.Runs[0].Text =
                "Edited header contents.";

            // When we compare documents, the differences of the latter document from the former show up as revisions to the former
            // Each edit that we've made above will have its own revision, after we run the Compare method
            // We can compare with a CompareOptions object, which can suppress changes done to certain types of objects within the original document
            // from registering as revisions after the comparison by setting some of these members to "true"
            CompareOptions compareOptions = new CompareOptions();
            compareOptions.IgnoreFormatting = false;
            compareOptions.IgnoreCaseChanges = false;
            compareOptions.IgnoreComments = false;
            compareOptions.IgnoreTables = false;
            compareOptions.IgnoreFields = false;
            compareOptions.IgnoreFootnotes = false;
            compareOptions.IgnoreTextboxes = false;
            compareOptions.IgnoreHeadersAndFooters = false;
            compareOptions.Target = ComparisonTargetType.New;

            docOriginal.Compare(docEdited, "John Doe", DateTime.Now, compareOptions);
            docOriginal.Save(ArtifactsDir + "Document.CompareOptions.docx");
            //ExEnd

            docOriginal = new Document(ArtifactsDir + "Document.CompareOptions.docx");

            TestUtil.VerifyFootnote(FootnoteType.Endnote, true, string.Empty,
                "OriginalEdited endnote text.", (Footnote)docOriginal.GetChild(NodeType.Footnote, 0, true));

            // If we set compareOptions to ignore certain types of changes,
            // then revisions done on those types of nodes will not appear in the output document
            // We can tell what kind of node a revision was done on by looking at the NodeType of the revision's parent nodes
            Assert.AreNotEqual(compareOptions.IgnoreFormatting,
                docOriginal.Revisions.Any(rev => rev.RevisionType == RevisionType.FormatChange));
            Assert.AreNotEqual(compareOptions.IgnoreCaseChanges,
                docOriginal.Revisions.Any(s => s.ParentNode.GetText().Contains("hello")));
            Assert.AreNotEqual(compareOptions.IgnoreComments,
                docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Comment)));
            Assert.AreNotEqual(compareOptions.IgnoreTables,
                docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Table)));
            Assert.AreNotEqual(compareOptions.IgnoreFields,
                docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.FieldStart)));
            Assert.AreNotEqual(compareOptions.IgnoreFootnotes,
                docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Footnote)));
            Assert.AreNotEqual(compareOptions.IgnoreTextboxes,
                docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.Shape)));
            Assert.AreNotEqual(compareOptions.IgnoreHeadersAndFooters,
                docOriginal.Revisions.Any(rev => HasParentOfType(rev, NodeType.HeaderFooter)));
        }

        /// <summary>
        /// Returns true if the passed revision has a parent node with the type specified by parentType
        /// </summary>
        private bool HasParentOfType(Revision revision, NodeType parentType)
        {
            Node n = revision.ParentNode;
            while (n.ParentNode != null)
            {
                if (n.NodeType == parentType) return true;
                n = n.ParentNode;
            }

            return false;
        }

        [Test]
        public void RemoveExternalSchemaReferences()
        {
            //ExStart
            //ExFor:Document.RemoveExternalSchemaReferences
            //ExSummary:Shows how to remove all external XML schema references from a document. 
            Document doc = new Document(MyDir + "External XML schema.docx");

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
            Document doc = new Document();
            Assert.AreEqual(4, doc.Styles.Count); //ExSkip

            // Insert some styles into a blank document
            doc.Styles.Add(StyleType.List, "MyListStyle1");
            doc.Styles.Add(StyleType.List, "MyListStyle2");
            doc.Styles.Add(StyleType.Character, "MyParagraphStyle1");
            doc.Styles.Add(StyleType.Character, "MyParagraphStyle2");

            // Combined with the built in styles, the document now has 8 styles in total,
            // but all 4 of the ones we added count as unused
            Assert.AreEqual(8, doc.Styles.Count);

            // A character style counts as used when the document contains text in that style
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Style = doc.Styles["MyParagraphStyle1"];
            builder.Writeln("Hello world!");

            // A list style is also "used" when there is a list that uses it
            Aspose.Words.Lists.List list = doc.Lists.Add(doc.Styles["MyListStyle1"]);
            builder.ListFormat.List = list;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            // The Cleanup() method, when configured with a CleanupOptions object, can target unused styles and remove them
            CleanupOptions cleanupOptions = new CleanupOptions();
            cleanupOptions.UnusedLists = true;
            cleanupOptions.UnusedStyles = true;
            
            // We've added 4 styles and used 2 of them, so the other two will be removed when this method is called
            doc.Cleanup(cleanupOptions);
            Assert.AreEqual(6, doc.Styles.Count);
            //ExEnd

            doc.FirstSection.Body.RemoveAllChildren();
            doc.Cleanup(cleanupOptions);

            Assert.AreEqual(4, doc.Styles.Count);
        }

        [Test]
        public void RemoveDuplicateStyles()
        {
            //ExStart
            //ExFor:CleanupOptions.DuplicateStyle
            //ExSummary:Shows how to remove duplicated styles from the document.
            Document doc = new Document(MyDir + "Document.docx");
            
            CleanupOptions options = new CleanupOptions();
            options.DuplicateStyle = true;
 
            doc.Cleanup(options);
            doc.Save(ArtifactsDir + "Document.RemoveDuplicateStyles.docx");
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

            // This text will appear as normal text in the document and no revisions will be counted
            doc.FirstSection.Body.FirstParagraph.Runs.Add(new Run(doc, "Hello world!"));
            Assert.AreEqual(0, doc.Revisions.Count);

            doc.StartTrackRevisions("Author");

            // This text will appear as a revision
            // We did not specify a time while calling StartTrackRevisions(), so the date/time that's noted
            // on the revision will be the real time when StartTrackRevisions() executes
            doc.FirstSection.Body.AppendParagraph("Hello again!");
            Assert.AreEqual(2, doc.Revisions.Count);

            // Stopping the tracking of revisions makes this text appear as normal text
            // Revisions are not counted when the document is changed
            doc.StopTrackRevisions();
            doc.FirstSection.Body.AppendParagraph("Hello again!");
            Assert.AreEqual(2, doc.Revisions.Count);

            // Specifying some date/time will apply that date/time to all subsequent revisions until StopTrackRevisions() is called
            // Note that placing values such as DateTime.MinValue as an argument will create revisions that do not have a date/time at all
            doc.StartTrackRevisions("Author", new DateTime(1970, 1, 1));
            doc.FirstSection.Body.AppendParagraph("Hello again!");
            Assert.AreEqual(4, doc.Revisions.Count);

            doc.Save(ArtifactsDir + "Document.StartTrackRevisions.docx");
            //ExEnd
        }

        [Test]
        public void ShowRevisionBalloons()
        {
            //ExStart
            //ExFor:RevisionOptions.ShowInBalloons
            //ExSummary:Shows how render tracking changes in balloons
            Document doc = new Document(MyDir + "Revisions.docx");

            // Set option true, if you need render tracking changes in balloons in pdf document,
            // while comments will stay visible
            doc.LayoutOptions.RevisionOptions.ShowInBalloons = ShowInBalloons.None;

            // Check that revisions are in balloons 
            doc.Save(ArtifactsDir + "Document.ShowRevisionBalloons.pdf");
            //ExEnd
        }

        [Test]
        public void AcceptAllRevisions()
        {
            //ExStart
            //ExFor:Document.AcceptAllRevisions
            //ExSummary:Shows how to accept all tracking changes in the document.
            Document doc = new Document(MyDir + "Document.docx");

            // Start tracking and make some revisions
            doc.StartTrackRevisions("Author");
            doc.FirstSection.Body.AppendParagraph("Hello world!");
            Assert.AreEqual(2, doc.Revisions.Count); //ExSkip

            // Revisions will now show up as normal text in the output document
            doc.AcceptAllRevisions();
            doc.Save(ArtifactsDir + "Document.AcceptAllRevisions.docx");
            Assert.AreEqual(0, doc.Revisions.Count); //ExSKip
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
            //ExFor:Story.Paragraphs
            //ExSummary:Shows how to get paragraph that was moved (deleted/inserted) in Microsoft Word while change tracking was enabled.
            Document doc = new Document(MyDir + "Revisions.docx");

            // There are two sets of move revisions in this document
            // One moves a small part of a paragraph, while the other moves a whole paragraph
            // Paragraph.IsMoveFromRevision/IsMoveToRevision will only be true if a whole paragraph is moved, as in the latter case
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].IsMoveFromRevision)
                    Console.WriteLine("The paragraph {0} has been moved (deleted).", i);
                if (paragraphs[i].IsMoveToRevision)
                    Console.WriteLine("The paragraph {0} has been moved (inserted).", i);
            }
            //ExEnd

            Assert.AreEqual(11, doc.Revisions.Count());
            Assert.AreEqual(6, doc.Revisions.Count(r => r.RevisionType == RevisionType.Moving));
            Assert.AreEqual(1, paragraphs.Count(p => ((Paragraph)p).IsMoveFromRevision));
            Assert.AreEqual(1, paragraphs.Count(p => ((Paragraph)p).IsMoveToRevision));
        }

        [Test]
        public void GetRevisedPropertiesOfList()
        {
            //ExStart
            //ExFor:RevisionsView
            //ExFor:Document.RevisionsView
            //ExSummary:Shows how to get revised version of list label and list level formatting in a document.
            Document doc = new Document(MyDir + "Revisions at list levels.docx");
            doc.UpdateListLabels();

            // Switch to the revised version of the document
            doc.RevisionsView = RevisionsView.Final;

            foreach (Revision revision in doc.Revisions)
            {
                if (revision.ParentNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph paragraph = (Paragraph)revision.ParentNode;

                    if (paragraph.IsListItem)
                    {
                        // Print revised version of LabelString and ListLevel
                        Console.WriteLine(paragraph.ListLabel.LabelString);
                        Console.WriteLine(paragraph.ListFormat.ListLevel);
                    }
                }
            }
            //ExEnd

            Assert.AreEqual("", ((Paragraph)doc.Revisions[0].ParentNode).ListLabel.LabelString);
            Assert.AreEqual("1.", ((Paragraph)doc.Revisions[1].ParentNode).ListLabel.LabelString);
            Assert.AreEqual("a.", ((Paragraph)doc.Revisions[3].ParentNode).ListLabel.LabelString);

            doc.RevisionsView = RevisionsView.Original;

            Assert.AreEqual("1.", ((Paragraph)doc.Revisions[0].ParentNode).ListLabel.LabelString);
            Assert.AreEqual("a.", ((Paragraph)doc.Revisions[1].ParentNode).ListLabel.LabelString);
            Assert.AreEqual("", ((Paragraph)doc.Revisions[3].ParentNode).ListLabel.LabelString);
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
            Document doc = new Document(MyDir + "Rendering.docx");

            // If we aren't setting the thumbnail via built in document properties,
            // we can set the first page of the document to be the thumbnail in an output .epub like this
            doc.UpdateThumbnail();
            doc.Save(ArtifactsDir + "Document.UpdateThumbnail.FirstPage.epub");

            // Another way is to use the first image shape found in the document as the thumbnail
            // Insert an image with a builder that we want to use as a thumbnail
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(ImageDir + "Logo.jpg");

            ThumbnailGeneratingOptions options = new ThumbnailGeneratingOptions();
            Assert.AreEqual(new Size(600, 900), options.ThumbnailSize); //ExSKip
            Assert.IsTrue(options.GenerateFromFirstPage); //ExSkip
            options.ThumbnailSize = new Size(400, 400);
            options.GenerateFromFirstPage = false;

            doc.UpdateThumbnail(options);
            doc.Save(ArtifactsDir + "Document.UpdateThumbnail.FirstImage.epub");
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

            doc.Save(ArtifactsDir + "Document.HyphenationOptions.docx");
            //ExEnd

            Assert.AreEqual(true, doc.HyphenationOptions.AutoHyphenation);
            Assert.AreEqual(2, doc.HyphenationOptions.ConsecutiveHyphenLimit);
            Assert.AreEqual(720, doc.HyphenationOptions.HyphenationZone);
            Assert.AreEqual(true, doc.HyphenationOptions.HyphenateCaps);

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "Document.HyphenationOptions.docx",
                GoldsDir + "Document.HyphenationOptions Gold.docx"));
        }

        [Test]
        public void HyphenationOptionsDefaultValues()
        {
            Document doc = new Document();
            doc = DocumentHelper.SaveOpen(doc);

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
            //ExSummary:Shows how to simply extract text from a document.
            TxtLoadOptions loadOptions = new TxtLoadOptions { DetectNumberingWithWhitespaces = false };

            PlainTextDocument plaintext = new PlainTextDocument(MyDir + "Document.docx");
            Assert.AreEqual("Hello World!", plaintext.Text.Trim()); //ExSkip 

            plaintext = new PlainTextDocument(MyDir + "Document.docx", loadOptions);
            Assert.AreEqual("Hello World!", plaintext.Text.Trim()); //ExSkip
            //ExEnd
        }

        [Test]
        public void GetPlainTextBuiltInDocumentProperties()
        {
            //ExStart
            //ExFor:PlainTextDocument.BuiltInDocumentProperties
            //ExSummary:Shows how to get BuiltIn properties of plain text document.
            PlainTextDocument plaintext = new PlainTextDocument(MyDir + "Bookmarks.docx");
            BuiltInDocumentProperties builtInDocumentProperties = plaintext.BuiltInDocumentProperties;
            //ExEnd

            Assert.AreEqual("Aspose", builtInDocumentProperties.Company);
        }

        [Test]
        public void GetPlainTextCustomDocumentProperties()
        {
            //ExStart
            //ExFor:PlainTextDocument.CustomDocumentProperties
            //ExSummary:Shows how to get custom properties of plain text document.
            PlainTextDocument plaintext = new PlainTextDocument(MyDir + "Bookmarks.docx");
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
            //ExSummary:Shows how to simply extract text from a stream.
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.DetectNumberingWithWhitespaces = false;

            using (Stream stream = new FileStream(MyDir + "Document.docx", FileMode.Open))
            {
                PlainTextDocument plaintext = new PlainTextDocument(stream);
                Assert.AreEqual("Hello World!", plaintext.Text.Trim()); //ExSkip

                plaintext = new PlainTextDocument(stream, loadOptions);
                Assert.AreEqual("Hello World!", plaintext.Text.Trim()); //ExSkip
            }
            //ExEnd
        }

        [Test]
        public void OoxmlComplianceVersion()
        {
            //ExStart
            //ExFor:Document.Compliance
            //ExSummary:Shows how to get OOXML compliance version.
            // Open a DOC and check its OOXML compliance version
            Document doc = new Document(MyDir + "Document.doc");

            OoxmlCompliance compliance = doc.Compliance;
            Assert.AreEqual(compliance, OoxmlCompliance.Ecma376_2006);

            // Open a DOCX which should have a newer one
            doc = new Document(MyDir + "Document.docx");
            compliance = doc.Compliance;

            Assert.AreEqual(compliance, OoxmlCompliance.Iso29500_2008_Transitional);
            //ExEnd
        }

        [Test, Ignore("WORDSNET-20342")]
        public void ImageSaveOptions()
        {
            //ExStart
            //ExFor:Document.Save(Stream, String, Saving.SaveOptions)
            //ExFor:SaveOptions.UseAntiAliasing
            //ExFor:SaveOptions.UseHighQualityRendering
            //ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
            Document doc = new Document(MyDir + "Rendering.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Size = 60;
            builder.Writeln("Some text.");

            SaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);
            Assert.IsFalse(options.UseAntiAliasing); //ExSkip
            Assert.IsFalse(options.UseHighQualityRendering); //ExSkip

            doc.Save(ArtifactsDir + "Document.ImageSaveOptions.Default.jpg", options);

            options.UseAntiAliasing = true;
            options.UseHighQualityRendering = true;

            doc.Save(ArtifactsDir + "Document.ImageSaveOptions.HighQuality.jpg", options);
            //ExEnd

            TestUtil.VerifyImage(794, 1122, ArtifactsDir + "Document.ImageSaveOptions.Default.jpg");
            TestUtil.VerifyImage(794, 1122, ArtifactsDir + "Document.ImageSaveOptions.HighQuality.jpg");
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

            // Add two styles and apply them to the builder's formats, marking them as "used" 
            builder.ParagraphFormat.Style = doc.Styles.Add(StyleType.Paragraph, "My Used Style");
            builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDiamonds);

            // And two more styles and leave them unused by not applying them to anything
            doc.Styles.Add(StyleType.Paragraph, "My Unused Style");
            doc.Lists.Add(ListTemplate.NumberArabicDot);
            Assert.NotNull(doc.Styles["My Used Style"]); //ExSkip
            Assert.NotNull(doc.Styles["My Unused Style"]); //ExSkip
            Assert.IsTrue(doc.Lists.Any(l => l.ListLevels[0].NumberStyle == NumberStyle.Bullet)); //ExSkip
            Assert.IsTrue(doc.Lists.Any(l => l.ListLevels[0].NumberStyle == NumberStyle.Arabic)); //ExSkip

            doc.Cleanup();

            // The used styles are still in the document
            Assert.NotNull(doc.Styles["My Used Style"]);
            Assert.IsTrue(doc.Lists.Any(l => l.ListLevels[0].NumberStyle == NumberStyle.Bullet));

            // The unused styles have been removed
            Assert.IsNull(doc.Styles["My Unused Style"]);
            Assert.IsFalse(doc.Lists.Any(l => l.ListLevels[0].NumberStyle == NumberStyle.Arabic));
            //ExEnd

            Assert.AreEqual(5, doc.Styles.Count); 
            Assert.AreEqual(1, doc.Lists.Count);

            doc.RemoveAllChildren();
            doc.Cleanup();

            Assert.AreEqual(4, doc.Styles.Count);
            Assert.AreEqual(0, doc.Lists.Count);
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
            //ExFor:RevisionType
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
            Document doc = new Document(MyDir + "Revisions.docx");
            RevisionCollection revisions = doc.Revisions;
            
            // This collection itself has a collection of revision groups, which are merged sequences of adjacent revisions
            Assert.AreEqual(7, revisions.Groups.Count); //ExSkip
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
            Assert.AreEqual(11, revisions.Count); //ExSkip
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
            //ExEnd
        }

        [Test]
        public void AutomaticallyUpdateStyles()
        {
            //ExStart
            //ExFor:Document.AutomaticallyUpdateStyles
            //ExSummary:Shows how to update a document's styles based on its template.
            Document doc = new Document();

            // Empty Microsoft Word documents by default come with an attached template called "Normal.dotm"
            // There is no default template for Aspose Words documents
            Assert.AreEqual(string.Empty, doc.AttachedTemplate);

            // For AutomaticallyUpdateStyles to have any effect, we need a document with a template
            // We can make a document with word and open it
            // Or we can attach a template from our file system, as below
            doc.AttachedTemplate = MyDir + "Business brochure.dotx";

            // Any changes to the styles in this template will be propagated to those styles in the document
            doc.AutomaticallyUpdateStyles = true;

            doc.Save(ArtifactsDir + "Document.AutomaticallyUpdateStyles.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.AutomaticallyUpdateStyles.docx");

            Assert.True(doc.AutomaticallyUpdateStyles);
            Assert.AreEqual(MyDir + "Business brochure.dotx", doc.AttachedTemplate);
            Assert.True(File.Exists(doc.AttachedTemplate));
        }

        [Test]
        public void DefaultTemplate()
        {
            //ExStart
            //ExFor:Document.AttachedTemplate
            //ExFor:SaveOptions.CreateSaveOptions(String)
            //ExFor:SaveOptions.DefaultTemplate
            //ExSummary:Shows how to set a default .docx document template.
            Document doc = new Document();

            // If we set this flag to true while not having a template attached to the document,
            // there will be no effect because there is no template document to draw style changes from
            doc.AutomaticallyUpdateStyles = true;
            Assert.That(doc.AttachedTemplate, Is.Empty);

            // We can set a default template document filename in a SaveOptions object to make it apply to
            // all documents we save with it that have no AttachedTemplate value
            SaveOptions options = SaveOptions.CreateSaveOptions("Document.DefaultTemplate.docx");
            options.DefaultTemplate = MyDir + "Business brochure.dotx";
            Assert.True(File.Exists(options.DefaultTemplate)); //ExSkip

            doc.Save(ArtifactsDir + "Document.DefaultTemplate.docx", options);
            //ExEnd
        }

        [Test]
        public void Sections()
        {
            //ExStart
            //ExFor:Document.LastSection
            //ExSummary:Shows how to edit the last section of a document.
            // Open the template document, containing obsolete copyright information in the footer
            Document doc = new Document(MyDir + "Footer.docx");
            
            // Create a new copyright information string to replace an older one with
            int currentYear = DateTime.Now.Year;
            string newCopyrightInformation = $"Copyright (C) {currentYear} by Aspose Pty Ltd.";
            
            FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
            findReplaceOptions.MatchCase = false;
            findReplaceOptions.FindWholeWordsOnly = false;
            
            // Each section has its own set of headers/footers,
            // so the text in each one has to be replaced individually if we want the entire document to be affected
            HeaderFooter firstSectionFooter = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
            firstSectionFooter.Range.Replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

            HeaderFooter lastSectionFooter = doc.LastSection.HeadersFooters[HeaderFooterType.FooterPrimary];
            lastSectionFooter.Range.Replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

            doc.Save(ArtifactsDir + "Document.Sections.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.Sections.docx");
            Assert.AreEqual(doc.FirstSection, doc.Sections[0]);
            Assert.AreEqual(doc.LastSection, doc.Sections[1]);

            Assert.True(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetText().Contains($"Copyright (C) {currentYear} by Aspose Pty Ltd."));
            Assert.True(doc.LastSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetText().Contains($"Copyright (C) {currentYear} by Aspose Pty Ltd."));
            Assert.False(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetText().Contains("(C) 2006 Aspose Pty Ltd."));
            Assert.False(doc.LastSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetText().Contains("(C) 2006 Aspose Pty Ltd."));
        }

        //ExStart
        //ExFor:FindReplaceOptions.UseLegacyOrder
        //ExSummary:Shows how to include text box analyzing, during replacing text.
        [TestCase(true)] //ExSkip
        [TestCase(false)] //ExSkip
        public void UseLegacyOrder(bool isUseLegacyOrder)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert 3 tags to appear in sequential order, the second of which will be inside a text box
            builder.Writeln("[tag 1]");
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 50);
            builder.Writeln("[tag 3]");

            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("[tag 2]");

            UseLegacyOrderReplacingCallback callback = new UseLegacyOrderReplacingCallback();     
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = callback;

            // Use this option if want to search text sequentially from top to bottom considering the text boxes
            options.UseLegacyOrder = isUseLegacyOrder;
 
            doc.Range.Replace(new Regex(@"\[(.*?)\]"), "", options);
            CheckUseLegacyOrderResults(isUseLegacyOrder, callback); //ExSkip

            foreach (string match in ((UseLegacyOrderReplacingCallback)options.ReplacingCallback).Matches)
                Console.WriteLine(match);
        }

        private class UseLegacyOrderReplacingCallback : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Matches.Add(e.Match.Value); 
                return ReplaceAction.Replace;
            }

            public List<string> Matches { get; } = new List<string>();
        }
        //ExEnd

        private static void CheckUseLegacyOrderResults(bool isUseLegacyOrder, UseLegacyOrderReplacingCallback callback)
        {
            Assert.AreEqual(
                isUseLegacyOrder
                    ? new List<string> { "[tag 1]", "[tag 2]", "[tag 3]" }
                    : new List<string> { "[tag 1]", "[tag 3]", "[tag 2]" }, callback.Matches);
        }

        [Test]
        public void SetInvalidateFieldTypes()
        {
            //ExStart
            //ExFor:Document.NormalizeFieldTypes
            //ExSummary:Shows how to get the keep a field's type up to date with its field code.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a date field
            Field field = builder.InsertField("DATE", null);

            // Based on the field code we entered above, the type of the field has been set to "FieldDate"
            Assert.AreEqual(FieldType.FieldDate, field.Type);

            // We can manually access the content of the field we added and change it
            Run fieldText = (Run)doc.FirstSection.Body.FirstParagraph.GetChildNodes(NodeType.Run, true)[0];
            Assert.AreEqual("DATE", fieldText.Text); //ExSkip
            fieldText.Text = "PAGE";

            // We changed the text to "PAGE" but the field's type property did not update accordingly
            Assert.AreEqual("PAGE", fieldText.GetText());
            Assert.AreEqual(FieldType.FieldDate, field.Type);
            Assert.AreEqual(FieldType.FieldDate, field.Start.FieldType); //ExSkip
            Assert.AreEqual(FieldType.FieldDate, field.Separator.FieldType); //ExSkip
            Assert.AreEqual(FieldType.FieldDate, field.End.FieldType); //ExSkip

            // After running this method the type of the field, as well as its FieldStart,
            // FieldSeparator and FieldEnd nodes to "FieldPage", which matches the text "PAGE"
            doc.NormalizeFieldTypes();

            Assert.AreEqual(FieldType.FieldPage, field.Type);
            Assert.AreEqual(FieldType.FieldPage, field.Start.FieldType); //ExSkip
            Assert.AreEqual(FieldType.FieldPage, field.Separator.FieldType); //ExSkip
            Assert.AreEqual(FieldType.FieldPage, field.End.FieldType); //ExSkip
            //ExEnd
        }

        [Test]
        public void LayoutOptions()
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
            LayoutOptions options = doc.LayoutOptions;
            Assert.IsFalse(options.ShowHiddenText); //ExSkip
            Assert.IsFalse(options.ShowParagraphMarks); //ExSkip

            // The appearance of revisions can be controlled from the layout options property
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            Assert.AreEqual(RevisionColor.ByAuthor, options.RevisionOptions.InsertedTextColor); //ExSkip
            Assert.True(options.RevisionOptions.ShowRevisionBars); //ExSkip
            options.RevisionOptions.InsertedTextColor = RevisionColor.BrightGreen;
            options.RevisionOptions.ShowRevisionBars = false;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln(
                "This is a revision. Normally the text is red with a bar to the left, but we made some changes to the revision options.");

            doc.StopTrackRevisions();

            // Layout options can be used to show hidden text too
            builder.Writeln("This text is not hidden.");
            builder.Font.Hidden = true;
            builder.Writeln(
                "This text is hidden. It will only show up in the output if we allow it to via doc.LayoutOptions.");

            options.ShowHiddenText = true;

            // This option is equivalent to enabling paragraph marks in Microsoft Word via Home > paragraph > Show Paragraph Marks,
            // and can be used to display these features in a .pdf
            options.ShowParagraphMarks = true;

            doc.Save(ArtifactsDir + "Document.LayoutOptions.pdf");
            //ExEnd
        }

        [Test]
        public void MailMergeSettings()
        {
            //ExStart
            //ExFor:Document.MailMergeSettings
            //ExFor:MailMergeCheckErrors
            //ExFor:MailMergeDataType
            //ExFor:MailMergeDestination
            //ExFor:MailMergeMainDocumentType
            //ExFor:MailMergeSettings
            //ExFor:MailMergeSettings.CheckErrors
            //ExFor:MailMergeSettings.Clone
            //ExFor:MailMergeSettings.Destination
            //ExFor:MailMergeSettings.DataType
            //ExFor:MailMergeSettings.DoNotSupressBlankLines
            //ExFor:MailMergeSettings.LinkToQuery
            //ExFor:MailMergeSettings.MainDocumentType
            //ExFor:MailMergeSettings.Odso
            //ExFor:MailMergeSettings.Query
            //ExFor:MailMergeSettings.ViewMergedData
            //ExFor:Odso
            //ExFor:Odso.Clone
            //ExFor:Odso.ColumnDelimiter
            //ExFor:Odso.DataSource
            //ExFor:Odso.DataSourceType
            //ExFor:Odso.FirstRowContainsColumnNames
            //ExFor:OdsoDataSourceType
            //ExSummary:Shows how to execute an Office Data Source Object mail merge with MailMergeSettings.
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
            string dataSrcFilename = ArtifactsDir + "Document.MailMergeSettings.DataSource.txt";

            File.WriteAllLines(dataSrcFilename, lines);

            // Set the data source, query and other things
            MailMergeSettings settings = doc.MailMergeSettings;
            settings.MainDocumentType = MailMergeMainDocumentType.MailingLabels;
            settings.CheckErrors = MailMergeCheckErrors.Simulate;
            settings.DataType = MailMergeDataType.Native;
            settings.DataSource = dataSrcFilename;
            settings.Query = "SELECT * FROM " + doc.MailMergeSettings.DataSource;
            settings.LinkToQuery = true;
            settings.ViewMergedData = true;

            Assert.AreEqual(MailMergeDestination.Default, settings.Destination);
            Assert.False(settings.DoNotSupressBlankLines);

            // Office Data Source Object settings
            Odso odso = settings.Odso;
            odso.DataSource = dataSrcFilename;
            odso.DataSourceType = OdsoDataSourceType.Text;
            odso.ColumnDelimiter = '|';
            odso.FirstRowContainsColumnNames = true;

            // ODSO/MailMergeSettings objects can also be cloned
            Assert.AreNotSame(odso, odso.Clone());
            Assert.AreNotSame(settings, settings.Clone());

            // The mail merge will be performed when this document is opened 
            doc.Save(ArtifactsDir + "Document.MailMergeSettings.docx");
            //ExEnd

            settings = new Document(ArtifactsDir + "Document.MailMergeSettings.docx").MailMergeSettings;

            Assert.AreEqual(MailMergeMainDocumentType.MailingLabels, settings.MainDocumentType);
            Assert.AreEqual(MailMergeCheckErrors.Simulate, settings.CheckErrors);
            Assert.AreEqual(MailMergeDataType.Native, settings.DataType);
            Assert.AreEqual(ArtifactsDir + "Document.MailMergeSettings.DataSource.txt", settings.DataSource);
            Assert.AreEqual("SELECT * FROM " + doc.MailMergeSettings.DataSource, settings.Query);
            Assert.True(settings.LinkToQuery);
            Assert.True(settings.ViewMergedData);

            odso = settings.Odso;
            Assert.AreEqual(ArtifactsDir + "Document.MailMergeSettings.DataSource.txt", odso.DataSource);
            Assert.AreEqual(OdsoDataSourceType.Text, odso.DataSourceType);
            Assert.AreEqual('|', odso.ColumnDelimiter);
            Assert.True(odso.FirstRowContainsColumnNames);

        }

        [Test]
        public void OdsoEmail()
        {
            //ExStart
            //ExFor:MailMergeSettings.ActiveRecord
            //ExFor:MailMergeSettings.AddressFieldName
            //ExFor:MailMergeSettings.ConnectString
            //ExFor:MailMergeSettings.MailAsAttachment
            //ExFor:MailMergeSettings.MailSubject
            //ExFor:MailMergeSettings.Clear
            //ExFor:Odso.TableName
            //ExFor:Odso.UdlConnectString
            //ExSummary:Shows how to execute a mail merge while connecting to an external data source.
            Document doc = new Document(MyDir + "Odso data.docx");
            TestOdsoEmail(doc); //ExSkip
            MailMergeSettings settings = doc.MailMergeSettings;

            Console.WriteLine($"Connection string:\n\t{settings.ConnectString}");
            Console.WriteLine($"Mail merge docs as attachment:\n\t{settings.MailAsAttachment}");
            Console.WriteLine($"Mail merge doc e-mail subject:\n\t{settings.MailSubject}");
            Console.WriteLine($"Column that contains e-mail addresses:\n\t{settings.AddressFieldName}");
            Console.WriteLine($"Active record:\n\t{settings.ActiveRecord}");
            
            Odso odso = settings.Odso;

            Console.WriteLine($"File will connect to data source located in:\n\t\"{odso.DataSource}\"");
            Console.WriteLine($"Source type:\n\t{odso.DataSourceType}");
            Console.WriteLine($"UDL connection string:\n\t{odso.UdlConnectString}");
            Console.WriteLine($"Table:\n\t{odso.TableName}");
            Console.WriteLine($"Query:\n\t{doc.MailMergeSettings.Query}");

            // We can clear the settings, which will take place during saving
            settings.Clear();

            doc.Save(ArtifactsDir + "Document.OdsoEmail.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.OdsoEmail.docx");
            Assert.That(doc.MailMergeSettings.ConnectString, Is.Empty);
        }

        private void TestOdsoEmail(Document doc)
        {
            MailMergeSettings settings = doc.MailMergeSettings;

            Assert.False(settings.MailAsAttachment);
            Assert.AreEqual("test subject", settings.MailSubject);
            Assert.AreEqual("Email_Address", settings.AddressFieldName);
            Assert.AreEqual(66, settings.ActiveRecord);
            Assert.AreEqual("SELECT * FROM `Contacts` ", settings.Query);

            Odso odso = settings.Odso;

            Assert.AreEqual(settings.ConnectString, odso.UdlConnectString);
            Assert.AreEqual("Personal Folders|", odso.DataSource);
            Assert.AreEqual(OdsoDataSourceType.Email, odso.DataSourceType);
            Assert.AreEqual("Contacts", odso.TableName);
        }

        [Test]
        public void MailingLabelMerge()
        {
            //ExStart
            //ExFor:MailMergeSettings.DataSource
            //ExFor:MailMergeSettings.HeaderSource
            //ExSummary:Shows how to execute a mail merge while drawing data from a header and a data file.
            // Create a mailing label merge header file, which will consist of a table with one row 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();
            builder.Write("FirstName");
            builder.InsertCell();
            builder.Write("LastName");
            builder.EndTable();

            doc.Save(ArtifactsDir + "Document.MailingLabelMerge.Header.docx");

            // Create a mailing label merge date file, which will consist of a table with one row and the same amount of columns as 
            // the header table, which will determine the names for these columns
            doc = new Document();
            builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();
            builder.Write("John");
            builder.InsertCell();
            builder.Write("Doe");
            builder.EndTable();

            doc.Save(ArtifactsDir + "Document.MailingLabelMerge.Data.docx");

            // Create a merge destination document with MERGEFIELDS that will accept data
            doc = new Document();
            builder = new DocumentBuilder(doc);

            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");

            // Configure settings to draw data and headers from other documents
            MailMergeSettings settings = doc.MailMergeSettings;

            // The "header" document contains column names for the data in the "data" document,
            // which will correspond to the names of our MERGEFIELDs
            settings.HeaderSource = ArtifactsDir + "Document.MailingLabelMerge.Header.docx";
            settings.DataSource = ArtifactsDir + "Document.MailingLabelMerge.Data.docx";

            // Configure the rest of the MailMergeSettings object
            settings.Query = "SELECT * FROM " + settings.DataSource;
            settings.MainDocumentType = MailMergeMainDocumentType.MailingLabels;
            settings.DataType = MailMergeDataType.TextFile;
            settings.LinkToQuery = true;
            settings.ViewMergedData = true;

            // The mail merge will be performed when this document is opened 
            doc.Save(ArtifactsDir + "Document.MailingLabelMerge.docx");
            //ExEnd

            Assert.AreEqual("FirstName\aLastName\a\a", 
                new Document(ArtifactsDir + "Document.MailingLabelMerge.Header.docx").
                    GetChild(NodeType.Table, 0, true).GetText().Trim());

            Assert.AreEqual("John\aDoe\a\a",
                new Document(ArtifactsDir + "Document.MailingLabelMerge.Data.docx").
                    GetChild(NodeType.Table, 0, true).GetText().Trim());

            doc = new Document(ArtifactsDir + "Document.MailingLabelMerge.docx");

            Assert.AreEqual(2, doc.Range.Fields.Count);

            settings = doc.MailMergeSettings;

            Assert.AreEqual(ArtifactsDir + "Document.MailingLabelMerge.Header.docx", settings.HeaderSource);
            Assert.AreEqual(ArtifactsDir + "Document.MailingLabelMerge.Data.docx", settings.DataSource);
            Assert.AreEqual("SELECT * FROM " + settings.DataSource, settings.Query);
            Assert.AreEqual(MailMergeMainDocumentType.MailingLabels, settings.MainDocumentType);
            Assert.AreEqual(MailMergeDataType.TextFile, settings.DataType);
            Assert.True(settings.LinkToQuery);
            Assert.True(settings.ViewMergedData);
        }

        [Test]
        public void OdsoFieldMapDataCollection()
        {
            //ExStart
            //ExFor:Odso.FieldMapDatas
            //ExFor:OdsoFieldMapData
            //ExFor:OdsoFieldMapData.Clone
            //ExFor:OdsoFieldMapData.Column
            //ExFor:OdsoFieldMapData.MappedName
            //ExFor:OdsoFieldMapData.Name
            //ExFor:OdsoFieldMapData.Type
            //ExFor:OdsoFieldMapDataCollection
            //ExFor:OdsoFieldMapDataCollection.Add(OdsoFieldMapData)
            //ExFor:OdsoFieldMapDataCollection.Clear
            //ExFor:OdsoFieldMapDataCollection.Count
            //ExFor:OdsoFieldMapDataCollection.GetEnumerator
            //ExFor:OdsoFieldMapDataCollection.Item(Int32)
            //ExFor:OdsoFieldMapDataCollection.RemoveAt(Int32)
            //ExFor:OdsoFieldMappingType
            //ExSummary:Shows how to access the collection of data that maps data source columns to merge fields.
            Document doc = new Document(MyDir + "Odso data.docx");

            // This collection defines how columns from an external data source will be mapped to predefined MERGEFIELD,
            // ADDRESSBLOCK and GREETINGLINE fields during a mail merge
            OdsoFieldMapDataCollection dataCollection = doc.MailMergeSettings.Odso.FieldMapDatas;
            Assert.AreEqual(30, dataCollection.Count);

            using (IEnumerator<OdsoFieldMapData> enumerator = dataCollection.GetEnumerator())
            {
                int index = 0;
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"Field map data index {index++}, type \"{enumerator.Current.Type}\":");

                    Console.WriteLine(
                        enumerator.Current.Type != OdsoFieldMappingType.Null
                            ? $"\tColumn \"{enumerator.Current.Name}\", number {enumerator.Current.Column} mapped to merge field \"{enumerator.Current.MappedName}\"."
                            : "\tNo valid column to field mapping data present.");
                }
            }

            // Elements of the collection can be cloned
            Assert.AreNotEqual(dataCollection[0], dataCollection[0].Clone());

            // The collection can have individual entries removed or be cleared like this
            dataCollection.RemoveAt(0);
            Assert.AreEqual(29, dataCollection.Count); //ExSkip
            dataCollection.Clear();
            Assert.AreEqual(0, dataCollection.Count); //ExSkip
            //ExEnd
        }

        [Test]
        public void OdsoRecipientDataCollection()
        {
            //ExStart
            //ExFor:Odso.RecipientDatas
            //ExFor:OdsoRecipientData
            //ExFor:OdsoRecipientData.Active
            //ExFor:OdsoRecipientData.Clone
            //ExFor:OdsoRecipientData.Column
            //ExFor:OdsoRecipientData.Hash
            //ExFor:OdsoRecipientData.UniqueTag
            //ExFor:OdsoRecipientDataCollection
            //ExFor:OdsoRecipientDataCollection.Add(OdsoRecipientData)
            //ExFor:OdsoRecipientDataCollection.Clear
            //ExFor:OdsoRecipientDataCollection.Count
            //ExFor:OdsoRecipientDataCollection.GetEnumerator
            //ExFor:OdsoRecipientDataCollection.Item(Int32)
            //ExFor:OdsoRecipientDataCollection.RemoveAt(Int32)
            //ExSummary:Shows how to access the collection of data that designates merge data source records to be excluded from a merge.
            Document doc = new Document(MyDir + "Odso data.docx");

            // Records in this collection that do not have the "Active" flag set to true will be excluded from the mail merge
            OdsoRecipientDataCollection dataCollection = doc.MailMergeSettings.Odso.RecipientDatas;

            Assert.AreEqual(70, dataCollection.Count);

            using (IEnumerator<OdsoRecipientData> enumerator = dataCollection.GetEnumerator())
            {
                int index = 0;
                while (enumerator.MoveNext())
                {
                    Console.WriteLine(
                        $"Odso recipient data index {index++} will {(enumerator.Current.Active ? "" : "not ")}be imported upon mail merge.");
                    Console.WriteLine($"\tColumn #{enumerator.Current.Column}");
                    Console.WriteLine($"\tHash code: {enumerator.Current.Hash}");
                    Console.WriteLine($"\tContents array length: {enumerator.Current.UniqueTag.Length}");
                }
            }

            // Elements of the collection can be cloned
            Assert.AreNotEqual(dataCollection[0], dataCollection[0].Clone());

            // The collection can have individual entries removed or be cleared like this
            dataCollection.RemoveAt(0);
            Assert.AreEqual(69, dataCollection.Count); //ExSkip
            dataCollection.Clear();
            Assert.AreEqual(0, dataCollection.Count); //ExSkip
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
            Document doc = new Document(MyDir + "Custom parts OOXML package.docx");

            // Clone the second part
            CustomPart clonedPart = doc.PackageCustomParts[1].Clone();

            // Add the clone to the collection
            doc.PackageCustomParts.Add(clonedPart);
            TestDocPackageCustomParts(doc.PackageCustomParts); //ExSkip

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
                    Console.WriteLine(enumerator.Current.IsExternal
                        ? "\tSourced from outside the document"
                        : $"\tSourced from within the document, length: {enumerator.Current.Data.Length} bytes");
                    index++;
                }
            }

            // The parts collection can have individual entries removed or be cleared like this
            doc.PackageCustomParts.RemoveAt(2);
            Assert.AreEqual(2, doc.PackageCustomParts.Count); //ExSkip
            doc.PackageCustomParts.Clear();
            Assert.AreEqual(0, doc.PackageCustomParts.Count); //ExSkip
            //ExEnd
        }

        private static void TestDocPackageCustomParts(CustomPartCollection parts)
        {
            Assert.AreEqual(3, parts.Count);

            Assert.AreEqual("/payload/payload_on_package.test", parts[0].Name); 
            Assert.AreEqual("mytest/somedata", parts[0].ContentType); 
            Assert.AreEqual("http://mytest.payload.internal", parts[0].RelationshipType); 
            Assert.AreEqual(false, parts[0].IsExternal); 
            Assert.AreEqual(18, parts[0].Data.Length); 

            // This part is external and its content is sourced from outside the document
            Assert.AreEqual("http://www.aspose.com/Images/aspose-logo.jpg", parts[1].Name); 
            Assert.AreEqual("", parts[1].ContentType); 
            Assert.AreEqual("http://mytest.payload.external", parts[1].RelationshipType); 
            Assert.AreEqual(true, parts[1].IsExternal); 
            Assert.AreEqual(0, parts[1].Data.Length); 

            Assert.AreEqual("http://www.aspose.com/Images/aspose-logo.jpg", parts[2].Name); 
            Assert.AreEqual("", parts[2].ContentType); 
            Assert.AreEqual("http://mytest.payload.external", parts[2].RelationshipType); 
            Assert.AreEqual(true, parts[2].IsExternal); 
            Assert.AreEqual(0, parts[2].Data.Length); 
        }

        [Test]
        public void ShadeFormData()
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
                "If gray form field shading is turned on, this is the text that will have a gray background.", 0);

            // We can turn the grey shading off so the bookmarked text will blend in with the other text
            doc.ShadeFormData = false;
            doc.Save(ArtifactsDir + "Document.ShadeFormData.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.ShadeFormData.docx");
            Assert.IsFalse(doc.ShadeFormData);
        }

        [Test]
        public void VersionsCount()
        {
            //ExStart
            //ExFor:Document.VersionsCount
            //ExSummary:Shows how to count how many previous versions a document has.
            // Document versions are not supported but we can open an older document that has them
            Document doc = new Document(MyDir + "Versions.doc");

            // We can use this property to see how many there are
            // If we save and open the document, they will be lost
            Assert.AreEqual(4, doc.VersionsCount);
            //ExEnd

            doc.Save(ArtifactsDir + "Document.VersionsCount.docx");      
            doc = new Document(ArtifactsDir + "Document.VersionsCount.docx");

            Assert.AreEqual(0, doc.VersionsCount);
        }

        [Test]
        public void WriteProtection()
        {
            //ExStart
            //ExFor:Document.WriteProtection
            //ExFor:WriteProtection
            //ExFor:WriteProtection.IsWriteProtected
            //ExFor:WriteProtection.ReadOnlyRecommended
            //ExFor:WriteProtection.SetPassword(String)
            //ExFor:WriteProtection.ValidatePassword(String)
            //ExSummary:Shows how to protect a document with a password.
            Document doc = new Document();
            Assert.IsFalse(doc.WriteProtection.IsWriteProtected); //ExSkip
            Assert.IsFalse(doc.WriteProtection.ReadOnlyRecommended); //ExSkip

            // Enter a password that's up to 15 characters long
            doc.WriteProtection.SetPassword("MyPassword");

            Assert.IsTrue(doc.WriteProtection.IsWriteProtected);
            Assert.IsTrue(doc.WriteProtection.ValidatePassword("MyPassword"));

            // This flag applies to RTF documents and will be ignored by Microsoft Word
            doc.WriteProtection.ReadOnlyRecommended = true;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Write protection does not prevent us from editing the document programmatically.");

            // Save the document
            // Without the password, we can only read this document in Microsoft Word
            // With the password, we can read and write
            doc.Save(ArtifactsDir + "Document.WriteProtection.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.WriteProtection.docx");

            Assert.IsTrue(doc.WriteProtection.IsWriteProtected);
            Assert.IsTrue(doc.WriteProtection.ReadOnlyRecommended);
            Assert.IsTrue(doc.WriteProtection.ValidatePassword("MyPassword"));
            Assert.IsFalse(doc.WriteProtection.ValidatePassword("wrongpassword"));

            builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln("Writing text in a protected document.");

            Assert.AreEqual("Write protection does not prevent us from editing the document programmatically." +
                            "\rWriting text in a protected document.", doc.GetText().Trim());
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

            Document doc = new Document(MyDir + "No default editing language.docx", loadOptions);

            int localeIdFarEast = doc.Styles.DefaultFont.LocaleIdFarEast;
            Console.WriteLine(localeIdFarEast == (int)EditingLanguage.Japanese
                ? "The document either has no any FarEast language set in defaults or it was set to Japanese originally."
                : "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
            //ExEnd

            Assert.AreEqual((int)EditingLanguage.Japanese, doc.Styles.DefaultFont.LocaleIdFarEast);

            doc = new Document(MyDir + "No default editing language.docx");

            Assert.AreEqual((int)EditingLanguage.EnglishUS, doc.Styles.DefaultFont.LocaleIdFarEast);
        }

        [Test]
        public void SetEditingLanguageAsDefault()
        {
            //ExStart
            //ExFor:LanguagePreferences.DefaultEditingLanguage
            //ExSummary:Shows how to set language as default
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguagePreferences.DefaultEditingLanguage = EditingLanguage.Russian;

            Document doc = new Document(MyDir + "No default editing language.docx", loadOptions);

            int localeId = doc.Styles.DefaultFont.LocaleId;
            Console.WriteLine(localeId == (int)EditingLanguage.Russian
                ? "The document either has no any language set in defaults or it was set to Russian originally."
                : "The document default language was set to another than Russian language originally, so it is not overridden.");
            //ExEnd

            Assert.AreEqual((int)EditingLanguage.Russian, doc.Styles.DefaultFont.LocaleId);
            
            doc = new Document(MyDir + "No default editing language.docx");
            
            Assert.AreEqual((int)EditingLanguage.EnglishUS, doc.Styles.DefaultFont.LocaleId);
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
            //ExSummary:Shows how to get info about a group of revisions in document.
            Document doc = new Document(MyDir + "Revisions.docx");
            
            Assert.AreEqual(7, doc.Revisions.Groups.Count);

            // Get info about all of revisions in document
            foreach (RevisionGroup group in doc.Revisions.Groups)
            {
                Console.WriteLine(
                    $"Revision author: {group.Author}; Revision type: {group.RevisionType} \n\tRevision text: {group.Text}");
            }
            //ExEnd
        }

        [Test]
        public void GetSpecificRevisionGroup()
        {
            //ExStart
            //ExFor:RevisionGroupCollection
            //ExFor:RevisionGroupCollection.Item(Int32)
            //ExSummary:Shows how to get a group of revisions in document.
            Document doc = new Document(MyDir + "Revisions.docx");

            // Get revision group by index
            RevisionGroup revisionGroup = doc.Revisions.Groups[0];
            //ExEnd

            // Check revision group details
            Assert.AreEqual(RevisionType.Deletion, revisionGroup.RevisionType);
            Assert.AreEqual("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ", 
                revisionGroup.Text);
        }

        [Test]
        public void RemovePersonalInformation()
        {
            //ExStart
            //ExFor:Document.RemovePersonalInformation
            //ExSummary:Shows how to get or set a flag to remove all user information upon saving the MS Word document.
            Document doc = new Document(MyDir + "Revisions.docx");
            // If flag sets to 'true' that MS Word will remove all user information from comments, revisions and
            // document properties upon saving the document. In MS Word 2013 and 2016 you can see this using
            // File -> Options -> Trust Center -> Trust Center Settings -> Privacy Options -> then the
            // checkbox "Remove personal information from file properties on save"
            doc.RemovePersonalInformation = true;
            
            // Personal information will not be removed at this time
            // This will happen when we open this document in Microsoft Word and save it manually
            // Once noticeable change will be the revisions losing their author names
            doc.Save(ArtifactsDir + "Document.RemovePersonalInformation.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.RemovePersonalInformation.docx");
            Assert.IsTrue(doc.RemovePersonalInformation);
        }

        [Test]
        public void HideComments()
        {
            //ExStart
            //ExFor:LayoutOptions.ShowComments
            //ExSummary:Shows how to show or hide comments in PDF document.
            Document doc = new Document(MyDir + "Comments.docx");
            doc.LayoutOptions.ShowComments = false;
            
            doc.Save(ArtifactsDir + "Document.HideComments.pdf");
            //ExEnd

            Assert.False(doc.LayoutOptions.ShowComments);
        }

        [Test]
        public void RevisionOptions()
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
            //ExSummary:Shows how to edit appearance of revisions.
            Document doc = new Document(MyDir + "Revisions.docx");

            // Get the RevisionOptions object that controls the appearance of revisions
            RevisionOptions revisionOptions = doc.LayoutOptions.RevisionOptions;

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

            // Get movement, deletion, formatting revisions and comments to show up in green balloons on the right side of the page
            revisionOptions.ShowInBalloons = ShowInBalloons.Format;
            revisionOptions.CommentColor = RevisionColor.BrightGreen;

            // These features are only applicable to formats such as .pdf or .jpg
            doc.Save(ArtifactsDir + "Document.RevisionOptions.pdf");
            //ExEnd
        }

        [Test]
        public void CopyTemplateStylesViaDocument()
        {
            //ExStart
            //ExFor:Document.CopyStylesFromTemplate(Document)
            //ExSummary:Shows how to copies styles from the template to a document via Document.
            Document template = new Document(MyDir + "Rendering.docx");
            Document target = new Document(MyDir + "Document.docx");

            Assert.AreEqual(18, template.Styles.Count); //ExSkip
            Assert.AreEqual(4, target.Styles.Count); //ExSkip

            target.CopyStylesFromTemplate(template);
            Assert.AreEqual(18, target.Styles.Count); //ExSkip
            //ExEnd
        }

        [Test]
        public void CopyTemplateStylesViaString()
        {
            //ExStart
            //ExFor:Document.CopyStylesFromTemplate(String)
            //ExSummary:Shows how to copies styles from the template to a document via string.
            Document target = new Document(MyDir + "Document.docx");
            Assert.AreEqual(4, target.Styles.Count); //ExSkip

            target.CopyStylesFromTemplate(MyDir + "Rendering.docx");
            Assert.AreEqual(18, target.Styles.Count); //ExSkip
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
                Console.WriteLine(
                    $"\tStarts on page {layoutCollector.GetStartPageIndex(node)}, ends on page {layoutCollector.GetEndPageIndex(node)}," +
                    $" spanning {layoutCollector.GetNumPagesSpanned(node)} pages.");
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
        //ExSummary:Shows ways of traversing a document's layout entities.
        [Test] //ExSkip
        public void LayoutEnumerator()
        {
            // Open a document that contains a variety of layout entities
            // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
            // They are defined visually by the rectangular space that they occupy in the document
            Document doc = new Document(MyDir + "Layout entities.docx");

            // Create an enumerator that can traverse these entities like a tree
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
            Assert.AreEqual(doc, layoutEnumerator.Document);

            layoutEnumerator.MoveParent(LayoutEntityType.Page); 
            Assert.AreEqual(LayoutEntityType.Page, layoutEnumerator.Type);
            Assert.Throws<InvalidOperationException>(() => Console.WriteLine(layoutEnumerator.Text));

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
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order.
        /// </summary>
        private static void TraverseLayoutForward(LayoutEnumerator layoutEnumerator, int depth)
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
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order.
        /// </summary>
        private static void TraverseLayoutBackward(LayoutEnumerator layoutEnumerator, int depth)
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
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order.
        /// </summary>
        private static void TraverseLayoutForwardLogical(LayoutEnumerator layoutEnumerator, int depth)
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
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order.
        /// </summary>
        private static void TraverseLayoutBackwardLogical(LayoutEnumerator layoutEnumerator, int depth)
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
        /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent.
        /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document.
        /// </summary>
        private static void PrintCurrentEntity(LayoutEnumerator layoutEnumerator, int indent)
        {
            string tabs = new string('\t', indent);

            Console.WriteLine(layoutEnumerator.Kind == string.Empty
                ? $"{tabs}-> Entity type: {layoutEnumerator.Type}"
                : $"{tabs}-> Entity type & kind: {layoutEnumerator.Type}, {layoutEnumerator.Kind}");

            // Only spans can contain text
            if (layoutEnumerator.Type == LayoutEntityType.Span)
                Console.WriteLine($"{tabs}   Span contents: \"{layoutEnumerator.Text}\"");

            RectangleF leRect = layoutEnumerator.Rectangle;
            Console.WriteLine($"{tabs}   Rectangle dimensions {leRect.Width}x{leRect.Height}, X={leRect.X} Y={leRect.Y}");
            Console.WriteLine($"{tabs}   Page {layoutEnumerator.PageIndex}");
        }
        //ExEnd

        [Test]
        public void AlwaysCompressMetafiles()
        {
            //ExStart
            //ExFor:DocSaveOptions.AlwaysCompressMetafiles
            //ExSummary:Shows how to change metafiles compression in a document while saving.
            // Open a document that contains a Microsoft Equation 3.0 mathematical formula
            Document doc = new Document(MyDir + "Microsoft equation object.docx");
            
            // Large metafiles are always compressed when exporting a document in Aspose.Words, but small metafiles are not
            // compressed for performance reason. Some other document editors, such as LibreOffice, cannot read uncompressed
            // metafiles. The following option 'AlwaysCompressMetafiles' was introduced to choose appropriate behavior
            DocSaveOptions saveOptions = new DocSaveOptions();
            // False - small metafiles are not compressed for performance reason
            saveOptions.AlwaysCompressMetafiles = false;
            
            doc.Save(ArtifactsDir + "Document.AlwaysCompressMetafiles.False.docx", saveOptions);

            // True - all metafiles are compressed regardless of its size
            saveOptions.AlwaysCompressMetafiles = true;

            doc.Save(ArtifactsDir + "Document.AlwaysCompressMetafiles.True.docx", saveOptions);

            Assert.True(new FileInfo(ArtifactsDir + "Document.AlwaysCompressMetafiles.True.docx").Length <
                        new FileInfo(ArtifactsDir + "Document.AlwaysCompressMetafiles.False.docx").Length);
            //ExEnd
        }

        [Test]
        public void CreateNewVbaProject()
        {
            //ExStart
            //ExFor:VbaProject.#ctor
            //ExFor:VbaProject.Name
            //ExFor:VbaModule.#ctor
            //ExFor:VbaModule.Name
            //ExFor:VbaModule.Type
            //ExFor:VbaModule.SourceCode
            //ExFor:VbaModuleCollection.Add(VbaModule)
            //ExFor:VbaModuleType
            //ExSummary:Shows how to create a VbaProject from a scratch for using macros.
            Document doc = new Document();

            // Create a new VBA project
            VbaProject project = new VbaProject();
            project.Name = "Aspose.Project";
            doc.VbaProject = project;

            // Create a new module and specify a macro source code
            VbaModule module = new VbaModule();
            module.Name = "Aspose.Module";
            // VbaModuleType values:
            // procedural module - A collection of subroutines and functions
            // ------
            // document module - A type of VBA project item that specifies a module for embedded macros and programmatic access
            // operations that are associated with a document
            // ------
            // class module - A module that contains the definition for a new object. Each instance of a class creates
            // a new object, and procedures that are defined in the module become properties and methods of the object
            // ------
            // designer module - A VBA module that extends the methods and properties of an ActiveX control that has been
            // registered with the project
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = "New source code";

            // Add module to the VBA project
            doc.VbaProject.Modules.Add(module);

            doc.Save(ArtifactsDir + "Document.CreateVBAMacros.docm");
            //ExEnd

            project = new Document(ArtifactsDir + "Document.CreateVBAMacros.docm").VbaProject;

            Assert.AreEqual("Aspose.Project", project.Name);

            VbaModuleCollection modules = doc.VbaProject.Modules;

            Assert.AreEqual(2, modules.Count);

            Assert.AreEqual("ThisDocument", modules[0].Name);
            Assert.AreEqual(VbaModuleType.DocumentModule, modules[0].Type);
            Assert.Null(modules[0].SourceCode);

            Assert.AreEqual("Aspose.Module", modules[1].Name);
            Assert.AreEqual(VbaModuleType.ProceduralModule, modules[1].Type);
            Assert.AreEqual("New source code", modules[1].SourceCode);
        }

        [Test]
        public void CloneVbaProject()
        {
            //ExStart
            //ExFor:VbaProject.Clone
            //ExFor:VbaModule.Clone
            //ExSummary:Shows how to deep clone VbaProject and VbaModule.
            Document doc = new Document(MyDir + "VBA project.docm");
            Document destDoc = new Document();

            // Clone VbaProject to the document
            VbaProject copyVbaProject = doc.VbaProject.Clone();
            destDoc.VbaProject = copyVbaProject;

            // In destination document we already have "Module1", because he was cloned with VbaProject
            // Therefore need to remove it before cloning
            VbaModule oldVbaModule = destDoc.VbaProject.Modules["Module1"];
            VbaModule copyVbaModule = doc.VbaProject.Modules["Module1"].Clone();
            destDoc.VbaProject.Modules.Remove(oldVbaModule);
            destDoc.VbaProject.Modules.Add(copyVbaModule);

            destDoc.Save(ArtifactsDir + "Document.CloneVbaProject.docm");
            //ExEnd

            VbaProject originalVbaProject = new Document(ArtifactsDir + "Document.CloneVbaProject.docm").VbaProject;

            Assert.AreEqual(copyVbaProject.Name, originalVbaProject.Name);
            Assert.AreEqual(copyVbaProject.CodePage, originalVbaProject.CodePage);
            Assert.AreEqual(copyVbaProject.IsSigned, originalVbaProject.IsSigned);
            Assert.AreEqual(copyVbaProject.Modules.Count, originalVbaProject.Modules.Count);

            for (int i = 0; i < originalVbaProject.Modules.Count; i++)
            {
                Assert.AreEqual(copyVbaProject.Modules[i].Name, originalVbaProject.Modules[i].Name);
                Assert.AreEqual(copyVbaProject.Modules[i].Type, originalVbaProject.Modules[i].Type);
                Assert.AreEqual(copyVbaProject.Modules[i].SourceCode, originalVbaProject.Modules[i].SourceCode);
            }
        }

        [Test]
        public void ReadMacrosFromExistingDocument()
        {
            //ExStart
            //ExFor:Document.VbaProject
            //ExFor:VbaModuleCollection
            //ExFor:VbaModuleCollection.Count
            //ExFor:VbaModuleCollection.Item(System.Int32)
            //ExFor:VbaModuleCollection.Item(System.String)
            //ExFor:VbaModuleCollection.Remove
            //ExFor:VbaModule
            //ExFor:VbaModule.Name
            //ExFor:VbaModule.SourceCode
            //ExFor:VbaProject
            //ExFor:VbaProject.Name
            //ExFor:VbaProject.Modules
            //ExFor:VbaProject.CodePage
            //ExFor:VbaProject.IsSigned
            //ExSummary:Shows how to get access to VBA project information in the document.
            Document doc = new Document(MyDir + "VBA project.docm");

            // A VBA project inside the document is defined as a collection of VBA modules
            VbaProject vbaProject = doc.VbaProject;
            Assert.True(vbaProject.IsSigned); //ExSkip
            Console.WriteLine(vbaProject.IsSigned
                ? $"Project name: {vbaProject.Name} signed; Project code page: {vbaProject.CodePage}; Modules count: {vbaProject.Modules.Count()}\n"
                : $"Project name: {vbaProject.Name} not signed; Project code page: {vbaProject.CodePage}; Modules count: {vbaProject.Modules.Count()}\n");

            VbaModuleCollection vbaModules = doc.VbaProject.Modules; 

            Assert.AreEqual(vbaModules.Count(), 3);

            foreach (VbaModule module in vbaModules)
                Console.WriteLine($"Module name: {module.Name};\nModule code:\n{module.SourceCode}\n");

            // Set new source code for VBA module
            // You can retrieve object by integer or by name
            vbaModules[0].SourceCode = "Your VBA code...";
            vbaModules["Module1"].SourceCode = "Your VBA code...";

            // Remove one of VbaModule from VbaModuleCollection
            vbaModules.Remove(vbaModules[2]);
            //ExEnd

            Assert.AreEqual("AsposeVBAtest", vbaProject.Name);
            Assert.AreEqual(2, vbaProject.Modules.Count());
            Assert.AreEqual(1251, vbaProject.CodePage);
            Assert.False(vbaProject.IsSigned);

            Assert.AreEqual("ThisDocument", vbaModules[0].Name);
            Assert.AreEqual("Your VBA code...", vbaModules[0].SourceCode);

            Assert.AreEqual("Module1", vbaModules[1].Name);
            Assert.AreEqual("Your VBA code...", vbaModules[1].SourceCode);
        }

        [Test]
        public void SaveOutputParameters()
        {
            //ExStart
            //ExFor:SaveOutputParameters
            //ExFor:SaveOutputParameters.ContentType
            //ExSummary:Shows how to verify Content-Type strings from save output parameters.
            Document doc = new Document(MyDir + "Document.docx");

            // Save the document as a .doc and check parameters
            SaveOutputParameters parameters = doc.Save(ArtifactsDir + "Document.SaveOutputParameters.doc");
            Assert.AreEqual("application/msword", parameters.ContentType);

            // A .docx or a .pdf will have different parameters
            parameters = doc.Save(ArtifactsDir + "Document.SaveOutputParameters.pdf");
            Assert.AreEqual("application/pdf", parameters.ContentType);
            //ExEnd
        }

        [Test]
        public void SubDocument()
        {
            //ExStart
            //ExFor:SubDocument
            //ExFor:SubDocument.NodeType
            //ExSummary:Shows how to access a master document's subdocument.
            Document doc = new Document(MyDir + "Master document.docx");

            NodeCollection subDocuments = doc.GetChildNodes(NodeType.SubDocument, true);
            Assert.AreEqual(1, subDocuments.Count); //ExSkip

            SubDocument subDocument = (SubDocument)subDocuments[0];

            // The SubDocument object itself does not contain the documents of the subdocument and only serves as a reference
            Assert.False(subDocument.IsComposite);
            //ExEnd
        }

        [Test]
        public void CreateWebExtension()
        {
            //ExStart
            //ExFor:BaseWebExtensionCollection`1.Add(`0)
            //ExFor:TaskPane
            //ExFor:TaskPane.DockState
            //ExFor:TaskPane.IsVisible
            //ExFor:TaskPane.Width
            //ExFor:TaskPane.IsLocked
            //ExFor:TaskPane.WebExtension
            //ExFor:TaskPane.Row
            //ExFor:WebExtension
            //ExFor:WebExtension.Reference
            //ExFor:WebExtension.Properties
            //ExFor:WebExtension.Bindings
            //ExFor:WebExtension.IsFrozen
            //ExFor:WebExtensionReference.Id
            //ExFor:WebExtensionReference.Version
            //ExFor:WebExtensionReference.StoreType
            //ExFor:WebExtensionReference.Store
            //ExFor:WebExtensionPropertyCollection
            //ExFor:WebExtensionBindingCollection
            //ExFor:WebExtensionProperty.#ctor(String, String)
            //ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
            //ExFor:WebExtensionStoreType
            //ExFor:WebExtensionBindingType
            //ExFor:TaskPaneDockState
            //ExFor:TaskPaneCollection
            //ExSummary:Shows how to create add-ins inside the document.
            Document doc = new Document();

            // Create taskpane with "MyScript" add-in which will be used by the document
            TaskPane myScriptTaskPane = new TaskPane();
            doc.WebExtensionTaskPanes.Add(myScriptTaskPane);

            // Define task pane location when the document opens
            myScriptTaskPane.DockState = TaskPaneDockState.Right;
            myScriptTaskPane.IsVisible = true;
            myScriptTaskPane.Width = 300;
            myScriptTaskPane.IsLocked = true;
            // Use this option if you have several task panes
            myScriptTaskPane.Row = 1;

            // Add "MyScript Math Sample" add-in which will be displayed inside task pane
            WebExtension webExtension = myScriptTaskPane.WebExtension;

            // Application Id from store
            webExtension.Reference.Id = "WA104380646";
            // The current version of the application used
            webExtension.Reference.Version = "1.0.0.0";
            // Type of marketplace
            webExtension.Reference.StoreType = WebExtensionStoreType.OMEX;
            // Marketplace based on your locale
            webExtension.Reference.Store = CultureInfo.CurrentCulture.Name;

            webExtension.Properties.Add(new WebExtensionProperty("MyScript", "MyScript Math Sample"));
            webExtension.Bindings.Add(new WebExtensionBinding("MyScript", WebExtensionBindingType.Text, "104380646"));

            // Use this option if you need to block web extension from any action
            webExtension.IsFrozen = false;

            doc.Save(ArtifactsDir + "Document.WebExtension.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.WebExtension.docx");
            myScriptTaskPane = doc.WebExtensionTaskPanes[0];

            Assert.AreEqual(TaskPaneDockState.Right, myScriptTaskPane.DockState);
            Assert.True(myScriptTaskPane.IsVisible);
            Assert.AreEqual(300.0d, myScriptTaskPane.Width);
            Assert.True(myScriptTaskPane.IsLocked);
            Assert.AreEqual(1, myScriptTaskPane.Row);
            webExtension = myScriptTaskPane.WebExtension;

            Assert.AreEqual("WA104380646", webExtension.Reference.Id);
            Assert.AreEqual("1.0.0.0", webExtension.Reference.Version);
            Assert.AreEqual(WebExtensionStoreType.OMEX, webExtension.Reference.StoreType);
            Assert.AreEqual(CultureInfo.CurrentCulture.Name, webExtension.Reference.Store);

            Assert.AreEqual("MyScript", webExtension.Properties[0].Name);
            Assert.AreEqual("MyScript Math Sample", webExtension.Properties[0].Value);

            Assert.AreEqual("MyScript", webExtension.Bindings[0].Id);
            Assert.AreEqual(WebExtensionBindingType.Text, webExtension.Bindings[0].BindingType);
            Assert.AreEqual("104380646", webExtension.Bindings[0].AppRef);

            Assert.False(webExtension.IsFrozen);
        }

        [Test]
        public void GetWebExtensionInfo()
        {
            //ExStart
            //ExFor:BaseWebExtensionCollection`1
            //ExFor:BaseWebExtensionCollection`1.Add(`0)
            //ExFor:BaseWebExtensionCollection`1.Clear
            //ExFor:BaseWebExtensionCollection`1.GetEnumerator
            //ExFor:BaseWebExtensionCollection`1.Remove(Int32)
            //ExFor:BaseWebExtensionCollection`1.Count
            //ExFor:BaseWebExtensionCollection`1.Item(Int32)
            //ExSummary:Shows how to work with web extension collections.
            Document doc = new Document(MyDir + "Web extension.docx");
            Assert.AreEqual(1, doc.WebExtensionTaskPanes.Count); //ExSkip

            // Add new taskpane to the collection
            TaskPane newTaskPane = new TaskPane();
            doc.WebExtensionTaskPanes.Add(newTaskPane);
            Assert.AreEqual(2, doc.WebExtensionTaskPanes.Count); //ExSkip

            // Enumerate all WebExtensionProperty in a collection
            WebExtensionPropertyCollection webExtensionPropertyCollection = doc.WebExtensionTaskPanes[0].WebExtension.Properties;
            using (IEnumerator<WebExtensionProperty> enumerator = webExtensionPropertyCollection.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    WebExtensionProperty webExtensionProperty = enumerator.Current;
                    Console.WriteLine($"Binding name: {webExtensionProperty.Name}; Binding value: {webExtensionProperty.Value}");
                }
            }

            // We can remove task panes one by one or clear the entire collection
            doc.WebExtensionTaskPanes.Remove(1);
            Assert.AreEqual(1, doc.WebExtensionTaskPanes.Count); //ExSkip
            doc.WebExtensionTaskPanes.Clear();
            Assert.AreEqual(0, doc.WebExtensionTaskPanes.Count); //ExSkip
            //ExEnd
		}

		[Test]
        public void EpubCover()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // When saving to .epub, some Microsoft Word document properties can be converted to .epub metadata
            doc.BuiltInDocumentProperties.Author = "John Doe";
            doc.BuiltInDocumentProperties.Title = "My Book Title";

            // The thumbnail we specify here can become the cover image
            byte[] image = File.ReadAllBytes(ImageDir + "Transparent background logo.png");
            doc.BuiltInDocumentProperties.Thumbnail = image;

            doc.Save(ArtifactsDir + "Document.EpubCover.epub");
        }

        [Test]
        public void WorkWithWatermark()
        {
            //ExStart
            //ExFor:Watermark.SetText(String)
            //ExFor:Watermark.SetText(String, TextWatermarkOptions)
            //ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
            //ExFor:Watermark.Remove
            //ExFor:TextWatermarkOptions.FontFamily
            //ExFor:TextWatermarkOptions.FontSize
            //ExFor:TextWatermarkOptions.Color
            //ExFor:TextWatermarkOptions.Layout
            //ExFor:TextWatermarkOptions.IsSemitrasparent
            //ExFor:ImageWatermarkOptions.Scale
            //ExFor:ImageWatermarkOptions.IsWashout
            //ExFor:WatermarkLayout
            //ExFor:WatermarkType
            //ExSummary:Shows how to create and remove watermarks in the document.
            Document doc = new Document();

            doc.Watermark.SetText("Aspose Watermark");
            
            TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
            textWatermarkOptions.FontFamily = "Arial";
            textWatermarkOptions.FontSize = 36;
            textWatermarkOptions.Color = Color.Black;
            textWatermarkOptions.Layout = WatermarkLayout.Horizontal;
            textWatermarkOptions.IsSemitrasparent = false;

            doc.Watermark.SetText("Aspose Watermark", textWatermarkOptions);

            ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
            imageWatermarkOptions.Scale = 5;
            imageWatermarkOptions.IsWashout = false;
            
#if NET462 || JAVA
            doc.Watermark.SetImage(Image.FromFile(ImageDir + "Logo.jpg"), imageWatermarkOptions);
#elif NETCOREAPP2_1 || __MOBILE__
            using (SKBitmap image = SKBitmap.Decode(ImageDir + "Logo.jpg"))
            {
                doc.Watermark.SetImage(image, imageWatermarkOptions);
            }
#endif
            if (doc.Watermark.Type == WatermarkType.Text)
                doc.Watermark.Remove();
            //ExEnd
        }

        [Test]
        public void HideGrammarErrors()
        {
            //ExStart
            //ExFor:Document.ShowGrammaticalErrors
            //ExFor:Document.ShowSpellingErrors
            //ExSummary:Shows how to hide grammar errors in the document.
            Document doc = new Document(MyDir + "Document.docx");
            
            doc.ShowGrammaticalErrors = true;
            doc.ShowSpellingErrors = false;
            
            doc.Save(ArtifactsDir + "Document.HideGrammarErrors.docx");
            //ExEnd
        }
    }
}