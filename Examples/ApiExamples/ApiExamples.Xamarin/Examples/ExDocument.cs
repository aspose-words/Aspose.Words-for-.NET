// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Fonts;
using Aspose.Words.Layout;
using Aspose.Words.Loading;
using Aspose.Words.Markup;
using Aspose.Words.Notes;
using Aspose.Words.Rendering;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Aspose.Words.Vba;
using Aspose.Words.WebExtensions;
using NUnit.Framework;
using CompareOptions = Aspose.Words.Comparing.CompareOptions;
using MemoryFontSource = Aspose.Words.Fonts.MemoryFontSource;
#if NET462 || NETCOREAPP2_1 || JAVA
using Aspose.Pdf.Text;
using Aspose.Words.Shaping.HarfBuzz;
#endif
#if NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif
#if NET462 || MAC || JAVA
using Aspose.Words.Loading;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDocument : ApiExampleBase
    {
        [Test]
        public void Constructor()
        {
            //ExStart
            //ExFor:Document.#ctor(Boolean)
            //ExFor:Document.#ctor(String,LoadOptions)
            //ExSummary:Shows how to create and load documents.
            // There are two ways of creating a Document object using Aspose.Words.
            // 1 -  Create a blank document:
            Document doc = new Document();

            // New Document objects by default come with the minimal set of nodes
            // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
            doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, "Hello world!"));

            // 2 -  Load a document that exists in the local file system:
            doc = new Document(MyDir + "Document.docx");

            // Loaded documents will have contents that we can access and edit.
            Assert.AreEqual("Hello World!", doc.FirstSection.Body.FirstParagraph.GetText().Trim());

            // Some operations that need to occur during loading, such as using a password to decrypt a document,
            // can be done by passing a LoadOptions object when loading the document.
            doc = new Document(MyDir + "Encrypted.docx", new LoadOptions("docPassword"));

            Assert.AreEqual("Test encrypted document.", doc.FirstSection.Body.FirstParagraph.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void LoadFromStream()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream)
            //ExSummary:Shows how to load a document using a stream.
            using (Stream stream = File.OpenRead(MyDir + "Document.docx"))
            {
                Document doc = new Document(stream);

                Assert.AreEqual("Hello World!\r\rHello Word!\r\r\rHello World!", doc.GetText().Trim());
            }
            //ExEnd
        }

        [Test]
        public void LoadFromWeb()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream)
            //ExSummary:Shows how to load a document from a URL.
            // Create a URL that points to a Microsoft Word document.
            const string url = "https://omextemplates.content.office.net/support/templates/en-us/tf16402488.dotx";

            // Download the document into a byte array, then load that array into a document using a memory stream.
            using (WebClient webClient = new WebClient())
            {
                byte[] dataBytes = webClient.DownloadData(url);

                using (MemoryStream byteStream = new MemoryStream(dataBytes))
                {
                    Document doc = new Document(byteStream);

                    // At this stage, we can read and edit the document's contents and then save it to the local file system.
                    Assert.AreEqual("Use this section to highlight your relevant passions, activities, and how you like to give back. " +
                                    "It’s good to include Leadership and volunteer experiences here. " +
                                    "Or show off important extras like publications, certifications, languages and more.",
                        doc.FirstSection.Body.Paragraphs[4].GetText().Trim());

                    doc.Save(ArtifactsDir + "Document.LoadFromWeb.docx");
                }
            }
            //ExEnd

            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, url);
        }

        [Test]
        public void ConvertToPdf()
        {
            //ExStart
            //ExFor:Document.#ctor(String)
            //ExFor:Document.Save(String)
            //ExSummary:Shows how to open a document and convert it to .PDF.
            Document doc = new Document(MyDir + "Document.docx");
            
            doc.Save(ArtifactsDir + "Document.ConvertToPdf.pdf");
            //ExEnd
        }

        [Test]
        public void SaveToImageStream()
        {
            //ExStart
            //ExFor:Document.Save(Stream, SaveFormat)
            //ExSummary:Shows how to save a document to an image via stream, and then read the image from that stream.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Times New Roman";
            builder.Font.Size = 24;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            builder.InsertImage(ImageDir + "Logo.jpg");

#if NET462 || JAVA
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Bmp);

                stream.Position = 0;

                // Read the stream back into an image.
                using (Image image = Image.FromStream(stream))
                {
                    Assert.AreEqual(ImageFormat.Bmp, image.RawFormat);
                    Assert.AreEqual(816, image.Width);
                    Assert.AreEqual(1056, image.Height);
                }
            }
#elif NETCOREAPP2_1 || __MOBILE__
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Bmp);

                stream.Position = 0;

                SKCodec codec = SKCodec.Create(stream);

                Assert.AreEqual(SKEncodedImageFormat.Bmp, codec.EncodedFormat);

                stream.Position = 0;

                using (SKBitmap image = SKBitmap.Decode(stream))
                {
                    Assert.AreEqual(816, image.Width);
                    Assert.AreEqual(1056, image.Height);
                }
            }
#endif
            //ExEnd
        }

#if NET462 || NETCOREAPP2_1 || JAVA
        [Test, Category("IgnoreOnJenkins"), Category("SkipMono")]
        public void OpenType()
        {
            //ExStart
            //ExFor:LayoutOptions.TextShaperFactory
            //ExSummary:Shows how to support OpenType features using the HarfBuzz text shaping engine.
            Document doc = new Document(MyDir + "OpenType text shaping.docx");

            // Aspose.Words can use externally provided text shaper objects,
            // which represent fonts and compute shaping information for text.
            // A text shaper factory is necessary for documents that use multiple fonts.
            // When the text shaper factory set, the layout uses OpenType features.
            // An Instance property returns a static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory.
            doc.LayoutOptions.TextShaperFactory = HarfBuzzTextShaperFactory.Instance;

            // Currently, text shaping is performing when exporting to PDF or XPS formats.
            doc.Save(ArtifactsDir + "Document.OpenType.pdf");
            //ExEnd
        }
#endif

        [Test]
        public void DetectPdfDocumentFormat()
        {
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Pdf Document.pdf");
            Assert.AreEqual(info.LoadFormat, LoadFormat.Pdf);
        }

        [Test]
        public void OpenPdfDocument()
        {
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            Assert.AreEqual(
                "Heading 1\rHeading 1.1.1.1 Heading 1.1.1.2\rHeading 1.1.1.1.1.1.1.1.1 Heading 1.1.1.1.1.1.1.1.2\u000c",
                doc.Range.Text);
        }

        [Test]
        public void OpenProtectedPdfDocument()
        {
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EncryptionDetails = new PdfEncryptionDetails("Aspose", null, PdfEncryptionAlgorithm.RC4_40);

            doc.Save(ArtifactsDir + "Document.PdfDocumentEncrypted.pdf", saveOptions);

            PdfLoadOptions loadOptions = new PdfLoadOptions();
            loadOptions.Password = "Aspose";
            loadOptions.LoadFormat = LoadFormat.Pdf;

            doc = new Document(ArtifactsDir + "Document.PdfDocumentEncrypted.pdf", loadOptions);
        }
        
        [Test]
        public void OpenFromStreamWithBaseUri()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream,LoadOptions)
            //ExFor:LoadOptions.#ctor
            //ExFor:LoadOptions.BaseUri
            //ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
            using (Stream stream = File.OpenRead(MyDir + "Document.html"))
            {
                // Pass the URI of the base folder while loading it
                // so that any images with relative URIs in the HTML document can be found.
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.BaseUri = ImageDir;

                Document doc = new Document(stream, loadOptions);

                // Verify that the first shape of the document contains a valid image.
                Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

                Assert.IsTrue(shape.IsImage);
                Assert.IsNotNull(shape.ImageData.ImageBytes);
                Assert.AreEqual(32.0, ConvertUtil.PointToPixel(shape.Width), 0.01);
                Assert.AreEqual(32.0, ConvertUtil.PointToPixel(shape.Height), 0.01);
            }
            //ExEnd
        }

        [Test, Ignore("Need to rework.")]
        public void InsertHtmlFromWebPage()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream, LoadOptions)
            //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
            //ExFor:LoadFormat
            //ExSummary:Shows how save a web page as a .docx file.
            const string url = "http://www.aspose.com/";

            using (WebClient client = new WebClient()) 
            { 
                using (MemoryStream stream = new MemoryStream(client.DownloadData(url)))
                {
                    // The URL is used again as a baseUri to ensure that any relative image paths are retrieved correctly.
                    LoadOptions options = new LoadOptions(LoadFormat.Html, "", url);

                    // Load the HTML document from stream and pass the LoadOptions object.
                    Document doc = new Document(stream, options);

                    // At this stage, we can read and edit the document's contents and then save it to the local file system.
                    Assert.AreEqual("File Format APIs", doc.FirstSection.Body.Paragraphs[1].Runs[0].GetText().Trim()); //ExSkip

                    doc.Save(ArtifactsDir + "Document.InsertHtmlFromWebPage.docx");
                }
            }
            //ExEnd

            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, url);
        }

        [Test]
        public void LoadEncrypted()
        {
            //ExStart
            //ExFor:Document.#ctor(Stream,LoadOptions)
            //ExFor:Document.#ctor(String,LoadOptions)
            //ExFor:LoadOptions
            //ExFor:LoadOptions.#ctor(String)
            //ExSummary:Shows how to load an encrypted Microsoft Word document.
            Document doc;

            // Aspose.Words throw an exception if we try to open an encrypted document without its password.
            Assert.Throws<IncorrectPasswordException>(() => doc = new Document(MyDir + "Encrypted.docx"));

            // When loading such a document, the password is passed to the document's constructor using a LoadOptions object.
            LoadOptions options = new LoadOptions("docPassword");

            // There are two ways of loading an encrypted document with a LoadOptions object.
            // 1 -  Load the document from the local file system by filename:
            doc = new Document(MyDir + "Encrypted.docx", options);
            Assert.AreEqual("Test encrypted document.", doc.GetText().Trim()); //ExSkip

            // 2 -  Load the document from a stream:
            using (Stream stream = File.OpenRead(MyDir + "Encrypted.docx"))
            {
                doc = new Document(stream, options);
                Assert.AreEqual("Test encrypted document.", doc.GetText().Trim()); //ExSkip
            }
            //ExEnd
        }

        [Test]
        public void TempFolder()
        {
            //ExStart
            //ExFor:LoadOptions.TempFolder
            //ExSummary:Shows how to load a document using temporary files.
            // Note that such an approach can reduce memory usage but degrades speed
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.TempFolder = @"C:\TempFolder\";
            
            // Ensure that the directory exists and load
            Directory.CreateDirectory(loadOptions.TempFolder);
             
            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            //ExEnd
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
        public void ConvertToEpub()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "Document.ConvertToEpub.epub");
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

                // Verify that the stream contains the document.
                Assert.AreEqual("Hello World!\r\rHello Word!\r\r\rHello World!", new Document(dstStream).GetText().Trim());
            }
            //ExEnd
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
        //ExSummary:Shows how customize node changing with a callback.
        [Test] //ExSkip
        public void FontChangeViaCallback()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the node changing callback to custom implementation,
            // then add/remove nodes to get it to generate a log.
            HandleNodeChangingFontChanger callback = new HandleNodeChangingFontChanger();
            doc.NodeChangingCallback = callback;

            builder.Writeln("Hello world!");
            builder.Writeln("Hello again!");
            builder.InsertField(" HYPERLINK \"https://www.google.com/\" ");
            builder.InsertShape(ShapeType.Rectangle, 300, 300);

            doc.Range.Fields[0].Remove();

            Console.WriteLine(callback.GetLog());
            TestFontChangeViaCallback(callback.GetLog()); //ExSkip
        }

        /// <summary>
        /// Logs the date and time of each node insertion and removal.
        /// Sets a custom font name/size for the text contents of Run nodes.
        /// </summary>
        public class HandleNodeChangingFontChanger : INodeChangingCallback
        {
            void INodeChangingCallback.NodeInserted(NodeChangingArgs args)
            {
                mLog.AppendLine($"\tType:\t{args.Node.NodeType}");
                mLog.AppendLine($"\tHash:\t{args.Node.GetHashCode()}");

                if (args.Node.NodeType == NodeType.Run)
                {
                    Aspose.Words.Font font = ((Run) args.Node).Font;
                    mLog.Append($"\tFont:\tChanged from \"{font.Name}\" {font.Size}pt");

                    font.Size = 24;
                    font.Name = "Arial";

                    mLog.AppendLine($" to \"{font.Name}\" {font.Size}pt");
                    mLog.AppendLine($"\tContents:\n\t\t\"{args.Node.GetText()}\"");
                }
            }

            void INodeChangingCallback.NodeInserting(NodeChangingArgs args)
            {
                mLog.AppendLine($"\n{DateTime.Now:dd/MM/yyyy HH:mm:ss:fff}\tNode insertion:");
            }

            void INodeChangingCallback.NodeRemoved(NodeChangingArgs args)
            {
                mLog.AppendLine($"\tType:\t{args.Node.NodeType}");
                mLog.AppendLine($"\tHash code:\t{args.Node.GetHashCode()}");
            }

            void INodeChangingCallback.NodeRemoving(NodeChangingArgs args)
            {
                mLog.AppendLine($"\n{DateTime.Now:dd/MM/yyyy HH:mm:ss:fff}\tNode removal:");
            }

            public string GetLog()
            {
                return mLog.ToString();
            }

            private readonly StringBuilder mLog = new StringBuilder();
        }
        //ExEnd

        private static void TestFontChangeViaCallback(string log)
        {
            Assert.AreEqual(10, Regex.Matches(log, "insertion").Count);
            Assert.AreEqual(5, Regex.Matches(log, "removal").Count);
        }

        [Test]
        public void AppendDocument()
        {
            //ExStart
            //ExFor:Document.AppendDocument(Document, ImportFormatMode)
            //ExSummary:Shows how to append a document to the end of another document.
            Document srcDoc = new Document();
            srcDoc.FirstSection.Body.AppendParagraph("Source document text. ");

            Document dstDoc = new Document();
            dstDoc.FirstSection.Body.AppendParagraph("Destination document text. ");

            // Append the source document to the destination document while preserving its formatting,
            // then save the source document to the local file system.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            Assert.AreEqual(2, dstDoc.Sections.Count); //ExSkip

            dstDoc.Save(ArtifactsDir + "Document.AppendDocument.docx");
            //ExEnd

            string outDocText = new Document(ArtifactsDir + "Document.AppendDocument.docx").GetText();

            Assert.True(outDocText.StartsWith(dstDoc.GetText()));
            Assert.True(outDocText.EndsWith(srcDoc.GetText()));
        }

        [Test]
        // The file path used below does not point to an existing file.
        public void AppendDocumentFromAutomation()
        {
            Document doc = new Document();
            
            // We should call this method to clear this document of any existing content.
            doc.RemoveAllChildren();

            const int recordCount = 5;
            for (int i = 1; i <= recordCount; i++)
            {
                Document srcDoc = new Document();

                Assert.That(() => srcDoc == new Document("C:\\DetailsList.doc"),
                    Throws.TypeOf<FileNotFoundException>());

                // Append the source document at the end of the destination document.
                doc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

                // Automation required you to insert a new section break at this point, however, in Aspose.Words we
                // do not need to do anything here as the appended document is imported as separate sections already

                // Unlink all headers/footers in this section from the previous section headers/footers
                // if this is the second document or above being appended.
                if (i > 1)
                    Assert.That(() => doc.Sections[i].HeadersFooters.LinkToPrevious(false),
                        Throws.TypeOf<NullReferenceException>());
            }
        }

        [Test]
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
            //ExSummary:Shows how to validate and display information about each signature in a document.
            Document doc = new Document(MyDir + "Digitally signed.docx");

            foreach (DigitalSignature signature in doc.DigitalSignatures)
            {
                Console.WriteLine($"{(signature.IsValid ? "Valid" : "Invalid")} signature: ");
                Console.WriteLine($"\tReason:\t{signature.Comments}"); 
                Console.WriteLine($"\tType:\t{signature.SignatureType}");
                Console.WriteLine($"\tSign time:\t{signature.SignTime}");
                Console.WriteLine($"\tSubject name:\t{signature.CertificateHolder.Certificate.SubjectName}");
                Console.WriteLine($"\tIssuer name:\t{signature.CertificateHolder.Certificate.IssuerName.Name}");
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
            // Verify that a document is not signed.
            Assert.False(FileFormatUtil.DetectFileFormat(MyDir + "Document.docx").HasDigitalSignature);

            // Create a CertificateHolder object from a PKCS12 file, which we will use to sign the document.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);

            // There are two ways of saving a signed copy of a document to the local file system:
            // 1 - Designate a document by a local system filename and save a signed copy at a location specified by another filename.
            DigitalSignatureUtil.Sign(MyDir + "Document.docx", ArtifactsDir + "Document.DigitalSignature.docx", 
                certificateHolder, new SignOptions() { SignTime = DateTime.Now } );

            Assert.True(FileFormatUtil.DetectFileFormat(ArtifactsDir + "Document.DigitalSignature.docx").HasDigitalSignature);

            // 2 - Take a document from a stream and save a signed copy to another stream.
            using (FileStream inDoc = new FileStream(MyDir + "Document.docx", FileMode.Open))
            {
                using (FileStream outDoc = new FileStream(ArtifactsDir + "Document.DigitalSignature.docx", FileMode.Create))
                {
                    DigitalSignatureUtil.Sign(inDoc, outDoc, certificateHolder);
                }
            }

            Assert.True(FileFormatUtil.DetectFileFormat(ArtifactsDir + "Document.DigitalSignature.docx").HasDigitalSignature);

            // Please verify that all of the document's digital signatures are valid and check their details.
            Document signedDoc = new Document(ArtifactsDir + "Document.DigitalSignature.docx");
            DigitalSignatureCollection digitalSignatureCollection = signedDoc.DigitalSignatures;

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
            //ExStart
            //ExFor:Document.AppendDocument(Document, ImportFormatMode)
            //ExSummary:Shows how to append all the documents in a folder to the end of a template document.
            Document dstDoc = new Document();

            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Template Document");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Some content here");
            Assert.AreEqual(5, dstDoc.Styles.Count); //ExSkip
            Assert.AreEqual(1, dstDoc.Sections.Count); //ExSkip

            // Append all unencrypted documents with the .doc extension
            // from our local file system directory to the base document.
            List<string> docFiles = Directory.GetFiles(MyDir, "*.doc").Where(item => item.EndsWith(".doc")).ToList();
            foreach (string fileName in docFiles)
            {
                FileFormatInfo info = FileFormatUtil.DetectFileFormat(fileName);
                if (info.IsEncrypted)
                    continue;

                Document srcDoc = new Document(fileName);
                dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
            }

            dstDoc.Save(ArtifactsDir + "Document.AppendAllDocumentsInFolder.doc");
            //ExEnd

            Assert.AreEqual(7, dstDoc.Styles.Count);
            Assert.AreEqual(9, dstDoc.Sections.Count);
        }

        [Test]
        public void JoinRunsWithSameFormatting()
        {
            //ExStart
            //ExFor:Document.JoinRunsWithSameFormatting
            //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
            // Open a document that contains adjacent runs of text with identical formatting,
            // which commonly occurs if we edit the same paragraph multiple times in Microsoft Word.
            Document doc = new Document(MyDir + "Rendering.docx");

            // If any number of these runs are adjacent with identical formatting,
            // then the document may be simplified.
            Assert.AreEqual(317, doc.GetChildNodes(NodeType.Run, true).Count);

            // Combine such runs with this method and verify the number of run joins that will take place.
            Assert.AreEqual(121, doc.JoinRunsWithSameFormatting());

            // The number of joins and the number of runs we have after the join
            // should add up the number of runs we had initially.
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
            //ExSummary:Shows how to set a custom interval for tab stop positions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set tab stops to appear every 72 points (1 inch).
            builder.Document.DefaultTabStop = 72;

            // Each tab character snaps the text after it to the next closest tab stop position.
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
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            // Cloning will produce a new document with the same contents as the original,
            // but with a unique copy of each of the original document's nodes.
            Document clone = doc.Clone();

            Assert.AreEqual(doc.FirstSection.Body.FirstParagraph.Runs[0].GetText(), 
                clone.FirstSection.Body.FirstParagraph.Runs[0].Text);
            Assert.AreNotEqual(doc.FirstSection.Body.FirstParagraph.Runs[0].GetHashCode(),
                clone.FirstSection.Body.FirstParagraph.Runs[0].GetHashCode());
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

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Field");

            // GetText will retrieve the visible text as well as field codes and special characters.
            Assert.AreEqual("\u0013MERGEFIELD Field\u0014«Field»\u0015\u000c", doc.GetText());

            // ToString will give us the document's appearance if saved to a passed save format.
            Assert.AreEqual("«Field»\r\n", doc.ToString(SaveFormat.Text));
            //ExEnd
        }

        [Test]
        public void DocumentByteArray()
        {
            Document doc = new Document(MyDir + "Document.docx");

            MemoryStream streamOut = new MemoryStream();
            doc.Save(streamOut, SaveFormat.Docx);

            byte[] docBytes = streamOut.ToArray();

            MemoryStream streamIn = new MemoryStream(docBytes);

            Document loadDoc = new Document(streamIn);
            Assert.AreEqual(doc.GetText(), loadDoc.GetText());
        }

        [Test]
        public void ProtectUnprotect()
        {
            //ExStart
            //ExFor:Document.Protect(ProtectionType,String)
            //ExFor:Document.ProtectionType
            //ExFor:Document.Unprotect
            //ExFor:Document.Unprotect(String)
            //ExSummary:Shows how to protect and unprotect a document.
            Document doc = new Document();
            doc.Protect(ProtectionType.ReadOnly, "password");

            Assert.AreEqual(ProtectionType.ReadOnly, doc.ProtectionType);

            // If we open this document with Microsoft Word intending to edit it,
            // we will need to apply the password to get through the protection.
            doc.Save(ArtifactsDir + "Document.Protect.docx");

            // Note that the protection only applies to Microsoft Word users opening our document.
            // We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
            Document protectedDoc = new Document(ArtifactsDir + "Document.Protect.docx");

            Assert.AreEqual(ProtectionType.ReadOnly, protectedDoc.ProtectionType);

            DocumentBuilder builder = new DocumentBuilder(protectedDoc);
            builder.Writeln("Text added to a protected document.");
            Assert.AreEqual("Text added to a protected document.", protectedDoc.Range.Text.Trim()); //ExSkip

            // There are two ways of removing protection from a document.
            // 1 - With no password:
            doc.Unprotect();

            Assert.AreEqual(ProtectionType.NoProtection, doc.ProtectionType);

            doc.Protect(ProtectionType.ReadOnly, "NewPassword");

            Assert.AreEqual(ProtectionType.ReadOnly, doc.ProtectionType);

            doc.Unprotect("WrongPassword");

            Assert.AreEqual(ProtectionType.ReadOnly, doc.ProtectionType);

            // 2 - With the correct password:
            doc.Unprotect("NewPassword");

            Assert.AreEqual(ProtectionType.NoProtection, doc.ProtectionType);
            //ExEnd
        }

        [Test]
        public void DocumentEnsureMinimum()
        {
            //ExStart
            //ExFor:Document.EnsureMinimum
            //ExSummary:Shows how to ensure that a document contains the minimal set of nodes required for editing its contents.
            // A newly created document contains one child Section, which includes one child Body and one child Paragraph.
            // We can edit the document body's contents by adding nodes such as Runs or inline Shapes to that paragraph.
            Document doc = new Document();
            NodeCollection nodes = doc.GetChildNodes(NodeType.Any, true);

            Assert.AreEqual(NodeType.Section, nodes[0].NodeType);
            Assert.AreEqual(doc, nodes[0].ParentNode);

            Assert.AreEqual(NodeType.Body, nodes[1].NodeType);
            Assert.AreEqual(nodes[0], nodes[1].ParentNode);

            Assert.AreEqual(NodeType.Paragraph, nodes[2].NodeType);
            Assert.AreEqual(nodes[1], nodes[2].ParentNode);

            // This is the minimal set of nodes that we need to be able to edit the document.
            // We will no longer be able to edit the document if we remove any of them.
            doc.RemoveAllChildren();

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Any, true).Count);

            // Call this method to make sure that the document has at least those three nodes so we can edit it again.
            doc.EnsureMinimum();

            Assert.AreEqual(NodeType.Section, nodes[0].NodeType);
            Assert.AreEqual(NodeType.Body, nodes[1].NodeType);
            Assert.AreEqual(NodeType.Paragraph, nodes[2].NodeType);

            ((Paragraph)nodes[2]).Runs.Add(new Run(doc, "Hello world!"));
            //ExEnd

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
        }

        [Test]
        public void RemoveMacrosFromDocument()
        {
            //ExStart
            //ExFor:Document.RemoveMacros
            //ExSummary:Shows how to remove all macros from a document.
            Document doc = new Document(MyDir + "Macro.docm");

            Assert.IsTrue(doc.HasMacros);
            Assert.AreEqual("Project", doc.VbaProject.Name);

            // Remove the document's VBA project, along with all its macros.
            doc.RemoveMacros();

            Assert.IsFalse(doc.HasMacros);
            Assert.Null(doc.VbaProject);
            //ExEnd
        }

        [Test]
        public void GetPageCount()
        {
            //ExStart
            //ExFor:Document.PageCount
            //ExSummary:Shows how to count the number of pages in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Page 3");

            // Verify the expected page count of the document.
            Assert.AreEqual(3, doc.PageCount);

            // Getting the PageCount property invoked the document's page layout to calculate the value.
            // This operation will not need to be re-done when rendering the document to a fixed page save format,
            // such as .pdf. So you can save some time, especially with more complex documents.
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
            
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.Write("Ut enim ad minim veniam, " +
                            "quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

            // Aspose.Words does not track document metrics like these in real time.
            Assert.AreEqual(0, doc.BuiltInDocumentProperties.Characters);
            Assert.AreEqual(0, doc.BuiltInDocumentProperties.Words);
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.Paragraphs);
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.Lines);

            // To get accurate values for three of these properties, we will need to update them manually.
            doc.UpdateWordCount();

            Assert.AreEqual(196, doc.BuiltInDocumentProperties.Characters);
            Assert.AreEqual(36, doc.BuiltInDocumentProperties.Words);
            Assert.AreEqual(2, doc.BuiltInDocumentProperties.Paragraphs);

            // For the line count, we will need to call a specific overload of the updating method.
            Assert.AreEqual(1, doc.BuiltInDocumentProperties.Lines);

            doc.UpdateWordCount(true);

            Assert.AreEqual(4, doc.BuiltInDocumentProperties.Lines);
            //ExEnd
        }

        [Test]
        public void TableStyleToDirectFormatting()
        {
            //ExStart
            //ExFor:CompositeNode.GetChild
            //ExFor:Document.ExpandTableStylesToDirectFormatting
            //ExSummary:Shows how to apply the properties of a table's style directly to the table's elements.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Hello world!");
            builder.EndTable();

            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle1");
            tableStyle.RowStripe = 3;
            tableStyle.CellSpacing = 5;
            tableStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite;
            tableStyle.Borders.Color = Color.Blue;
            tableStyle.Borders.LineStyle = LineStyle.DotDash;

            table.Style = tableStyle;

            // This method concerns table style properties such as the ones we set above.
            doc.ExpandTableStylesToDirectFormatting();

            doc.Save(ArtifactsDir + "Document.TableStyleToDirectFormatting.docx");
            //ExEnd

            TestUtil.DocPackageFileContainsString("<w:tblStyleRowBandSize w:val=\"3\" />", 
                ArtifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml");
            TestUtil.DocPackageFileContainsString("<w:tblCellSpacing w:w=\"100\" w:type=\"dxa\" />", 
                ArtifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml");
            TestUtil.DocPackageFileContainsString("<w:tblBorders><w:top w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:left w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:bottom w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:right w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideH w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideV w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /></w:tblBorders>",
                ArtifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml");
        }

        [Test]
        public void UpdateTableLayout()
        {
            //ExStart
            //ExFor:Document.UpdateTableLayout
            //ExSummary:Shows how to preserve a table's layout when saving to .txt.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.EndTable();

            // Use a TxtSaveOptions object to preserve the table's layout when converting the document to plaintext.
            TxtSaveOptions options = new TxtSaveOptions();
            options.PreserveTableLayout = true;

            // Previewing the appearance of the document in .txt form shows that the table will not be represented accurately.
            Assert.AreEqual(0.0d, table.FirstRow.Cells[0].CellFormat.Width);
            Assert.AreEqual("CCC\r\neee\r\nlll\r\nlll\r\n   \r\n123\r\n\r\n", doc.ToString(options));

            // We can call UpdateTableLayout() to fix some of these issues.
            doc.UpdateTableLayout();

            Assert.AreEqual("Cell 1             Cell 2             Cell 3\r\n\r\n", doc.ToString(options));
            Assert.AreEqual(155.0d, table.FirstRow.Cells[0].CellFormat.Width, 2f);
            //ExEnd
        }

        [Test]
        public void GetOriginalFileInfo()
        {
            //ExStart
            //ExFor:Document.OriginalFileName
            //ExFor:Document.OriginalLoadFormat
            //ExSummary:Shows how to retrieve details of a document's load operation.
            Document doc = new Document(MyDir + "Document.docx");

            Assert.AreEqual(MyDir + "Document.docx", doc.OriginalFileName);
            Assert.AreEqual(LoadFormat.Docx, doc.OriginalLoadFormat);
            //ExEnd
        }

        [Test]
        [Description("WORDSNET-16099")]
        public void FootnoteColumns()
        {
            //ExStart
            //ExFor:FootnoteOptions
            //ExFor:FootnoteOptions.Columns
            //ExSummary:Shows how to split the footnote section into a given number of columns.
            Document doc = new Document(MyDir + "Footnotes and endnotes.docx");
            Assert.AreEqual(0, doc.FootnoteOptions.Columns); //ExSkip

            doc.FootnoteOptions.Columns = 2;
            doc.Save(ArtifactsDir + "Document.FootnoteColumns.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.FootnoteColumns.docx");

            Assert.AreEqual(2, doc.FirstSection.PageSetup.FootnoteOptions.Columns);
        }
        
        [Test]
        public void Compare()
        {
            //ExStart
            //ExFor:Document.Compare(Document, String, DateTime)
            //ExFor:RevisionCollection.AcceptAll
            //ExSummary:Shows how to compare documents. 
            Document docOriginal = new Document();
            DocumentBuilder builder = new DocumentBuilder(docOriginal);
            builder.Writeln("This is the original document.");

            Document docEdited = new Document();
            builder = new DocumentBuilder(docEdited);
            builder.Writeln("This is the edited document.");

            // Comparing documents with revisions will throw an exception.
            if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
                docOriginal.Compare(docEdited, "authorName", DateTime.Now);

            // After the comparison, the original document will gain a new revision
            // for every element that is different in the edited document.
            Assert.AreEqual(2, docOriginal.Revisions.Count); //ExSkip
            foreach (Revision r in docOriginal.Revisions)
            {
                Console.WriteLine($"Revision type: {r.RevisionType}, on a node of type \"{r.ParentNode.NodeType}\"");
                Console.WriteLine($"\tChanged text: \"{r.ParentNode.GetText()}\"");
            }

            // Accepting these revisions will transform the original document into the edited document.
            docOriginal.Revisions.AcceptAll();

            Assert.AreEqual(docOriginal.GetText(), docEdited.GetText());
            //ExEnd

            docOriginal = DocumentHelper.SaveOpen(docOriginal);
            Assert.AreEqual(0, docOriginal.Revisions.Count);
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
            //ExSummary:Shows how to filter specific types of document elements when making a comparison.
            // Create the original document and populate it with various kinds of elements.
            Document docOriginal = new Document();
            DocumentBuilder builder = new DocumentBuilder(docOriginal);

            // Paragraph text referenced with an endnote:
            builder.Writeln("Hello world! This is the first paragraph.");
            builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

            // Table:
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Original cell 1 text");
            builder.InsertCell();
            builder.Write("Original cell 2 text");
            builder.EndTable();

            // Textbox:
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 20);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Original textbox contents");

            // DATE field:
            builder.MoveTo(docOriginal.FirstSection.Body.AppendParagraph(""));
            builder.InsertField(" DATE ");

            // Comment:
            Comment newComment = new Comment(docOriginal, "John Doe", "J.D.", DateTime.Now);
            newComment.SetText("Original comment.");
            builder.CurrentParagraph.AppendChild(newComment);

            // Header:
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Original header contents.");

            // Create a clone of our document and perform a quick edit on each of the cloned document's elements.
            Document docEdited = (Document)docOriginal.Clone(true);
            Paragraph firstParagraph = docEdited.FirstSection.Body.FirstParagraph;

            firstParagraph.Runs[0].Text = "hello world! this is the first paragraph, after editing.";
            firstParagraph.ParagraphFormat.Style = docEdited.Styles[StyleIdentifier.Heading1];
            ((Footnote)docEdited.GetChild(NodeType.Footnote, 0, true)).FirstParagraph.Runs[1].Text = "Edited endnote text.";
            ((Table)docEdited.GetChild(NodeType.Table, 0, true)).FirstRow.Cells[1].FirstParagraph.Runs[0].Text = "Edited Cell 2 contents";
            ((Shape)docEdited.GetChild(NodeType.Shape, 0, true)).FirstParagraph.Runs[0].Text = "Edited textbox contents";
            ((FieldDate)docEdited.Range.Fields[0]).UseLunarCalendar = true; 
            ((Comment)docEdited.GetChild(NodeType.Comment, 0, true)).FirstParagraph.Runs[0].Text = "Edited comment.";
            docEdited.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].FirstParagraph.Runs[0].Text =
                "Edited header contents.";

            // Comparing documents creates a revision for every edit in the edited document.
            // A CompareOptions object has a series of flags that can suppress revisions
            // on each respective type of element, effectively ignoring their change.
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
            // then revisions done on those types of nodes will not appear in the output document.
            // We can tell what kind of node a revision was done by looking at the NodeType of the revision's parent nodes.
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
        /// Returns true if the passed revision has a parent node with the type specified by parentType.
        /// </summary>
        private static bool HasParentOfType(Revision revision, NodeType parentType)
        {
            Node n = revision.ParentNode;
            while (n.ParentNode != null)
            {
                if (n.NodeType == parentType) return true;
                n = n.ParentNode;
            }

            return false;
        }

        [TestCase(false)]
        [TestCase(true)]
        public void IgnoreDmlUniqueId(bool isIgnoreDmlUniqueId)
        {
            //ExStart
            //ExFor:CompareOptions.IgnoreDmlUniqueId
            //ExSummary:Shows how to compare documents ignoring DML unique ID.
            Document docA = new Document(MyDir + "DML unique ID original.docx");
            Document docB = new Document(MyDir + "DML unique ID compare.docx");
 
            // By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
            // If we are ignoring DML's unique ID, and revisions count were 0.
            CompareOptions compareOptions = new CompareOptions();
            compareOptions.IgnoreDmlUniqueId = isIgnoreDmlUniqueId;
 
            docA.Compare(docB, "Aspose.Words", DateTime.Now, compareOptions);

            Assert.AreEqual(isIgnoreDmlUniqueId ? 0 : 2, docA.Revisions.Count);
            //ExEnd
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
        public void TrackRevisions()
        {
            //ExStart
            //ExFor:Document.StartTrackRevisions(String)
            //ExFor:Document.StartTrackRevisions(String, DateTime)
            //ExFor:Document.StopTrackRevisions
            //ExSummary:Shows how to track revisions while editing a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Editing a document usually does not count as a revision until we begin tracking them.
            builder.Write("Hello world! ");

            Assert.AreEqual(0, doc.Revisions.Count);
            Assert.False(doc.FirstSection.Body.Paragraphs[0].Runs[0].IsInsertRevision);

            doc.StartTrackRevisions("John Doe");

            builder.Write("Hello again! ");

            Assert.AreEqual(1, doc.Revisions.Count);
            Assert.True(doc.FirstSection.Body.Paragraphs[0].Runs[1].IsInsertRevision);
            Assert.AreEqual("John Doe", doc.Revisions[0].Author);
            Assert.That(doc.Revisions[0].DateTime, Is.EqualTo(DateTime.Now).Within(10).Milliseconds);

            // Stop tracking revisions to not count any future edits as revisions.
            doc.StopTrackRevisions();
            builder.Write("Hello again! ");

            Assert.AreEqual(1, doc.Revisions.Count);
            Assert.False(doc.FirstSection.Body.Paragraphs[0].Runs[2].IsInsertRevision);

            // Creating revisions gives them a date and time of the operation.
            // We can disable this by passing DateTime.MinValue when we start tracking revisions.
            doc.StartTrackRevisions("John Doe", DateTime.MinValue);
            builder.Write("Hello again! ");

            Assert.AreEqual(2, doc.Revisions.Count);
            Assert.AreEqual("John Doe", doc.Revisions[1].Author);
            Assert.AreEqual(DateTime.MinValue, doc.Revisions[1].DateTime);
            
            // We can accept/reject these revisions programmatically
            // by calling methods such as Document.AcceptAllRevisions, or each revision's Accept method.
            // In Microsoft Word, we can process them manually via "Review" -> "Changes".
            doc.Save(ArtifactsDir + "Document.StartTrackRevisions.docx");
            //ExEnd
        }
        
        [Test]
        public void AcceptAllRevisions()
        {
            //ExStart
            //ExFor:Document.AcceptAllRevisions
            //ExSummary:Shows how to accept all tracking changes in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Edit the document while tracking changes to create a few revisions.
            doc.StartTrackRevisions("John Doe");
            builder.Write("Hello world! ");
            builder.Write("Hello again! "); 
            builder.Write("This is another revision.");
            doc.StopTrackRevisions();

            Assert.AreEqual(3, doc.Revisions.Count);

            // We can iterate through every revision and accept/reject it as a part of our document.
            // If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
            doc.AcceptAllRevisions();

            Assert.AreEqual(0, doc.Revisions.Count);
            Assert.AreEqual("Hello world! Hello again! This is another revision.", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void GetRevisedPropertiesOfList()
        {
            //ExStart
            //ExFor:RevisionsView
            //ExFor:Document.RevisionsView
            //ExSummary:Shows how to switch between the revised and the original view of a document.
            Document doc = new Document(MyDir + "Revisions at list levels.docx");
            doc.UpdateListLabels();

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            Assert.AreEqual("1.", paragraphs[0].ListLabel.LabelString);
            Assert.AreEqual("a.", paragraphs[1].ListLabel.LabelString);
            Assert.AreEqual(string.Empty, paragraphs[2].ListLabel.LabelString);

            // View the document object as if all the revisions are accepted. Currently supports list labels.
            doc.RevisionsView = RevisionsView.Final;

            Assert.AreEqual(string.Empty, paragraphs[0].ListLabel.LabelString);
            Assert.AreEqual("1.", paragraphs[1].ListLabel.LabelString);
            Assert.AreEqual("a.", paragraphs[2].ListLabel.LabelString);
            //ExEnd

            doc.RevisionsView = RevisionsView.Original;
            doc.AcceptAllRevisions();

            Assert.AreEqual("a.", paragraphs[0].ListLabel.LabelString);
            Assert.AreEqual(string.Empty, paragraphs[1].ListLabel.LabelString);
            Assert.AreEqual("b.", paragraphs[2].ListLabel.LabelString);
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
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            builder.InsertImage(ImageDir + "Logo.jpg");

            // There are two ways of setting a thumbnail image when saving a document to .epub.
            // 1 -  Use the document's first page:
            doc.UpdateThumbnail();
            doc.Save(ArtifactsDir + "Document.UpdateThumbnail.FirstPage.epub");

            // 2 -  Use the first image found in the document:
            ThumbnailGeneratingOptions options = new ThumbnailGeneratingOptions();
            Assert.AreEqual(new Size(600, 900), options.ThumbnailSize); //ExSKip
            Assert.True(options.GenerateFromFirstPage); //ExSkip
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
            //ExSummary:Shows how to configure automatic hyphenation.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Size = 24;
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720;
            doc.HyphenationOptions.HyphenateCaps = true;

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
        public void OoxmlComplianceVersion()
        {
            //ExStart
            //ExFor:Document.Compliance
            //ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
            // The compliance version varies between documents created by different versions of Microsoft Word.
            Document doc = new Document(MyDir + "Document.doc");

            Assert.AreEqual(doc.Compliance, OoxmlCompliance.Ecma376_2006);

            doc = new Document(MyDir + "Document.docx");

            Assert.AreEqual(doc.Compliance, OoxmlCompliance.Iso29500_2008_Transitional);
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
        public void Cleanup()
        {
            //ExStart
            //ExFor:Document.Cleanup
            //ExSummary:Shows how to remove unused custom styles from a document.
            Document doc = new Document();

            doc.Styles.Add(StyleType.List, "MyListStyle1");
            doc.Styles.Add(StyleType.List, "MyListStyle2");
            doc.Styles.Add(StyleType.Character, "MyParagraphStyle1");
            doc.Styles.Add(StyleType.Character, "MyParagraphStyle2");

            // Combined with the built-in styles, the document now has eight styles.
            // A custom style counts as "used" while applied to some part of the document,
            // which means that the four styles we added are currently unused.
            Assert.AreEqual(8, doc.Styles.Count);

            // Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Style = doc.Styles["MyParagraphStyle1"];
            builder.Writeln("Hello world!");

            Aspose.Words.Lists.List list = doc.Lists.Add(doc.Styles["MyListStyle1"]);
            builder.ListFormat.List = list;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            doc.Cleanup();

            Assert.AreEqual(6, doc.Styles.Count);

            // Removing every node that a custom style is applied to marks it as "unused" again.
            // Run the Cleanup method again to remove them.
            doc.FirstSection.Body.RemoveAllChildren();
            doc.Cleanup();

            Assert.AreEqual(4, doc.Styles.Count);
            //ExEnd
        }
        
        [Test]
        public void AutomaticallyUpdateStyles()
        {
            //ExStart
            //ExFor:Document.AutomaticallyUpdateStyles
            //ExSummary:Shows how to attach a template to a document.
            Document doc = new Document();

            // Microsoft Word documents by default come with an attached template called "Normal.dotm".
            // There is no default template for blank Aspose.Words documents.
            Assert.AreEqual(string.Empty, doc.AttachedTemplate);

            // Attach a template, then set the flag to apply style changes
            // within the template to styles in our document.
            doc.AttachedTemplate = MyDir + "Business brochure.dotx";
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
            //ExFor:Document.AutomaticallyUpdateStyles
            //ExFor:SaveOptions.CreateSaveOptions(String)
            //ExFor:SaveOptions.DefaultTemplate
            //ExSummary:Shows how to set a default template for documents that do not have attached templates.
            Document doc = new Document();

            // Enable automatic style updating, but do not attach a template document.
            doc.AutomaticallyUpdateStyles = true;

            Assert.AreEqual(string.Empty, doc.AttachedTemplate);

            // Since there is no template document, the document had nowhere to track style changes.
            // Use a SaveOptions object to automatically set a template
            // if a document that we are saving does not have one.
            SaveOptions options = SaveOptions.CreateSaveOptions("Document.DefaultTemplate.docx");
            options.DefaultTemplate = MyDir + "Business brochure.dotx";

            doc.Save(ArtifactsDir + "Document.DefaultTemplate.docx", options);
            //ExEnd

            Assert.True(File.Exists(options.DefaultTemplate));
        }

        [Test]
        public void UseSubstitutions()
        {
            //ExStart
            //ExFor:FindReplaceOptions.UseSubstitutions
            //ExFor:FindReplaceOptions.LegacyMode
            //ExSummary:Shows how to recognize and use substitutions within replacement patterns.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
             
            builder.Write("Jason gave money to Paul.");
             
            Regex regex = new Regex(@"([A-z]+) gave money to ([A-z]+)");
             
            FindReplaceOptions options = new FindReplaceOptions();
            options.UseSubstitutions = true;

            // Using legacy mode does not support many advanced features, so we need to set it to 'false'.
            options.LegacyMode = false;

            doc.Range.Replace(regex, @"$2 took money from $1", options);
            
            Assert.AreEqual(doc.GetText(), "Paul took money from Jason.\f");
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

            Field field = builder.InsertField("DATE", null);

            // Aspose.Words automatically detects field types based on field codes.
            Assert.AreEqual(FieldType.FieldDate, field.Type);

            // Manually change the raw text of the field, which determines the field code.
            Run fieldText = (Run)doc.FirstSection.Body.FirstParagraph.GetChildNodes(NodeType.Run, true)[0];
            Assert.AreEqual("DATE", fieldText.Text); //ExSkip
            fieldText.Text = "PAGE";

            // Changing the field code has changed this field to one of a different type,
            // but the field's type properties still display the old type.
            Assert.AreEqual("PAGE", field.GetFieldCode());
            Assert.AreEqual(FieldType.FieldDate, field.Type);
            Assert.AreEqual(FieldType.FieldDate, field.Start.FieldType);
            Assert.AreEqual(FieldType.FieldDate, field.Separator.FieldType);
            Assert.AreEqual(FieldType.FieldDate, field.End.FieldType);

            // Update those properties with this method to display current value.
            doc.NormalizeFieldTypes();

            Assert.AreEqual(FieldType.FieldPage, field.Type);
            Assert.AreEqual(FieldType.FieldPage, field.Start.FieldType);
            Assert.AreEqual(FieldType.FieldPage, field.Separator.FieldType); 
            Assert.AreEqual(FieldType.FieldPage, field.End.FieldType);
            //ExEnd
        }

        [Test]
        public void LayoutOptionsRevisions()
        {
            //ExStart
            //ExFor:Document.LayoutOptions
            //ExFor:LayoutOptions
            //ExFor:LayoutOptions.RevisionOptions
            //ExFor:RevisionColor
            //ExFor:RevisionOptions
            //ExFor:RevisionOptions.InsertedTextColor
            //ExFor:RevisionOptions.ShowRevisionBars
            //ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a revision, then change the color of all revisions to green.
            builder.Writeln("This is not a revision.");
            doc.StartTrackRevisions("John Doe", DateTime.Now);
            Assert.AreEqual(RevisionColor.ByAuthor, doc.LayoutOptions.RevisionOptions.InsertedTextColor); //ExSkip
            Assert.True(doc.LayoutOptions.RevisionOptions.ShowRevisionBars); //ExSkip
            builder.Writeln("This is a revision.");
            doc.StopTrackRevisions();
            builder.Writeln("This is not a revision.");

            // Remove the bar that appears to the left of every revised line.
            doc.LayoutOptions.RevisionOptions.InsertedTextColor = RevisionColor.BrightGreen;
            doc.LayoutOptions.RevisionOptions.ShowRevisionBars = false;

            doc.Save(ArtifactsDir + "Document.LayoutOptionsRevisions.pdf");
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void LayoutOptionsHiddenText(bool showHiddenText)
        {
            //ExStart
            //ExFor:Document.LayoutOptions
            //ExFor:LayoutOptions
            //ExFor:Layout.LayoutOptions.ShowHiddenText
            //ExSummary:Shows how to hide text in a rendered output document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Assert.IsFalse(doc.LayoutOptions.ShowHiddenText); //ExSkip
            
            // Insert hidden text, then specify whether we wish to omit it from a rendered document.
            builder.Writeln("This text is not hidden.");
            builder.Font.Hidden = true;
            builder.Writeln("This text is hidden.");

            doc.LayoutOptions.ShowHiddenText = showHiddenText;

            doc.Save(ArtifactsDir + "Document.LayoutOptionsHiddenText.pdf");
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.LayoutOptionsHiddenText.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.AreEqual(showHiddenText ? 
                    $"This text is not hidden.{Environment.NewLine}This text is hidden." : 
                    "This text is not hidden.", textAbsorber.Text);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void LayoutOptionsParagraphMarks(bool showParagraphMarks)
        {
            //ExStart
            //ExFor:Document.LayoutOptions
            //ExFor:LayoutOptions
            //ExFor:Layout.LayoutOptions.ShowParagraphMarks
            //ExSummary:Shows how to show paragraph marks in a rendered output document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Assert.IsFalse(doc.LayoutOptions.ShowParagraphMarks); //ExSkip

            // Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
            // with a pilcrow (¶) symbol when we render the document.
            builder.Writeln("Hello world!");
            builder.Writeln("Hello again!");

            doc.LayoutOptions.ShowParagraphMarks = showParagraphMarks;

            doc.Save(ArtifactsDir + "Document.LayoutOptionsParagraphMarks.pdf");
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.LayoutOptionsParagraphMarks.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.AreEqual(showParagraphMarks ? 
                    $"Hello world!¶{Environment.NewLine}Hello again!¶{Environment.NewLine}¶" : 
                    $"Hello world!{Environment.NewLine}Hello again!", textAbsorber.Text);
#endif
        }

        [Test]
        public void UpdatePageLayout()
        {
            //ExStart
            //ExFor:StyleCollection.Item(String)
            //ExFor:SectionCollection.Item(Int32)
            //ExFor:Document.UpdatePageLayout
            //ExSummary:Shows when to recalculate the page layout of the document.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Saving a document to PDF, to an image, or printing for the first time will automatically
            // cache the layout of the document within its pages.
            doc.Save(ArtifactsDir + "Document.UpdatePageLayout.1.pdf");

            // Modify the document in some way.
            doc.Styles["Normal"].Font.Size = 6;
            doc.Sections[0].PageSetup.Orientation = Aspose.Words.Orientation.Landscape;

            // In the current version of Aspose.Words, modifying the document does not automatically rebuild 
            // the cached page layout. If we wish for the cached layout
            // to stay up to date, we will need to update it manually.
            doc.UpdatePageLayout();

            doc.Save(ArtifactsDir + "Document.UpdatePageLayout.2.pdf");
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
            //ExSummary:Shows how to access a document's arbitrary custom parts collection.
            Document doc = new Document(MyDir + "Custom parts OOXML package.docx");

            Assert.AreEqual(2, doc.PackageCustomParts.Count);

            // Clone the second part, then add the clone to the collection.
            CustomPart clonedPart = doc.PackageCustomParts[1].Clone();
            doc.PackageCustomParts.Add(clonedPart);
            TestDocPackageCustomParts(doc.PackageCustomParts); //ExSkip

            Assert.AreEqual(3, doc.PackageCustomParts.Count);

            // Enumerate over the collection and print every part.
            using (IEnumerator<CustomPart> enumerator = doc.PackageCustomParts.GetEnumerator())
            {
                int index = 0;
                while (enumerator.MoveNext())
                {
                    Console.WriteLine($"Part index {index}:");
                    Console.WriteLine($"\tName:\t\t\t\t{enumerator.Current.Name}");
                    Console.WriteLine($"\tContent type:\t\t{enumerator.Current.ContentType}");
                    Console.WriteLine($"\tRelationship type:\t{enumerator.Current.RelationshipType}");
                    Console.WriteLine(enumerator.Current.IsExternal ?
                        "\tSourced from outside the document" :
                        $"\tStored within the document, length: {enumerator.Current.Data.Length} bytes");
                    index++;
                }
            }

            // We can remove elements from this collection individually, or all at once.
            doc.PackageCustomParts.RemoveAt(2);

            Assert.AreEqual(2, doc.PackageCustomParts.Count);

            doc.PackageCustomParts.Clear();

            Assert.AreEqual(0, doc.PackageCustomParts.Count);
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

        [TestCase(false)]
        [TestCase(true)]
        public void ShadeFormData(bool useGreyShading)
        {
            //ExStart
            //ExFor:Document.ShadeFormData
            //ExSummary:Shows how to apply gray shading to form fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Assert.True(doc.ShadeFormData); //ExSkip

            builder.Write("Hello world! ");
            builder.InsertTextInput("My form field", TextFormFieldType.Regular, "",
                "Text contents of form field, which are shaded in grey by default.", 0);

            // We can turn the grey shading off, so the bookmarked text will blend in with the other text.
            doc.ShadeFormData = useGreyShading;
            doc.Save(ArtifactsDir + "Document.ShadeFormData.docx");
            //ExEnd
        }

        [Test]
        public void VersionsCount()
        {
            //ExStart
            //ExFor:Document.VersionsCount
            //ExSummary:Shows how to work with the versions count feature of older Microsoft Word documents.
            Document doc = new Document(MyDir + "Versions.doc");

            // We can read this property of a document, but we cannot preserve it while saving.
            Assert.AreEqual(4, doc.VersionsCount);

            doc.Save(ArtifactsDir + "Document.VersionsCount.doc");      
            doc = new Document(ArtifactsDir + "Document.VersionsCount.doc");

            Assert.AreEqual(0, doc.VersionsCount);
            //ExEnd
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
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world! This document is protected.");
            Assert.IsFalse(doc.WriteProtection.IsWriteProtected); //ExSkip
            Assert.IsFalse(doc.WriteProtection.ReadOnlyRecommended); //ExSkip

            // Enter a password up to 15 characters in length, and then verify the document's protection status.
            doc.WriteProtection.SetPassword("MyPassword");
            doc.WriteProtection.ReadOnlyRecommended = true;

            Assert.IsTrue(doc.WriteProtection.IsWriteProtected);
            Assert.IsTrue(doc.WriteProtection.ValidatePassword("MyPassword"));

            // Protection does not prevent the document from being edited programmatically, nor does it encrypt the contents.
            doc.Save(ArtifactsDir + "Document.WriteProtection.docx");
            doc = new Document(ArtifactsDir + "Document.WriteProtection.docx");

            Assert.IsTrue(doc.WriteProtection.IsWriteProtected);

            builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln("Writing text in a protected document.");

            Assert.AreEqual("Hello world! This document is protected." +
                            "\rWriting text in a protected document.", doc.GetText().Trim());
            //ExEnd
            Assert.IsTrue(doc.WriteProtection.ReadOnlyRecommended);
            Assert.IsTrue(doc.WriteProtection.ValidatePassword("MyPassword"));
            Assert.IsFalse(doc.WriteProtection.ValidatePassword("wrongpassword"));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void RemovePersonalInformation(bool saveWithoutPersonalInfo)
        {
            //ExStart
            //ExFor:Document.RemovePersonalInformation
            //ExSummary:Shows how to enable the removal of personal information during a manual save.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some content with personal information.
            doc.BuiltInDocumentProperties.Author = "John Doe";
            doc.BuiltInDocumentProperties.Company = "Placeholder Inc.";

            doc.StartTrackRevisions(doc.BuiltInDocumentProperties.Author, DateTime.Now);
            builder.Write("Hello world!");
            doc.StopTrackRevisions();

            // This flag is equivalent to File -> Options -> Trust Center -> Trust Center Settings... ->
            // Privacy Options -> "Remove personal information from file properties on save" in Microsoft Word.
            doc.RemovePersonalInformation = saveWithoutPersonalInfo;
            
            // This option will not take effect during a save operation made using Aspose.Words.
            // Personal data will be removed from our document with the flag set when we save it manually using Microsoft Word.
            doc.Save(ArtifactsDir + "Document.RemovePersonalInformation.docx");
            doc = new Document(ArtifactsDir + "Document.RemovePersonalInformation.docx");

            Assert.AreEqual(saveWithoutPersonalInfo, doc.RemovePersonalInformation);
            Assert.AreEqual("John Doe", doc.BuiltInDocumentProperties.Author);
            Assert.AreEqual("Placeholder Inc.", doc.BuiltInDocumentProperties.Company);
            Assert.AreEqual("John Doe", doc.Revisions[0].Author);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ShowComments(bool showComments)
        {
            //ExStart
            //ExFor:LayoutOptions.ShowComments
            //ExSummary:Shows how to show/hide comments when saving a document to a rendered format.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            comment.SetText("My comment.");
            builder.CurrentParagraph.AppendChild(comment);

            doc.LayoutOptions.ShowComments = showComments;

            doc.Save(ArtifactsDir + "Document.ShowComments.pdf");
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.ShowComments.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.AreEqual(
                showComments
                    ? "Hello world!                                                                    Commented [J.D.1]:  My comment."
                    : "Hello world!", textAbsorber.Text);
#endif
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
            Assert.AreEqual(8, target.Styles.Count); //ExSkip

            target.CopyStylesFromTemplate(template);
            Assert.AreEqual(18, target.Styles.Count); //ExSkip
            //ExEnd
        }

        [Test]
        public void CopyTemplateStylesViaDocumentNew()
        {
            //ExStart
            //ExFor:Document.CopyStylesFromTemplate(Document)
            //ExFor:Document.CopyStylesFromTemplate(String)
            //ExSummary:Shows how to copy styles from one document to another.
            // Create a document, and then add styles that we will copy to another document.
            Document template = new Document();
            
            Style style = template.Styles.Add(StyleType.Paragraph, "TemplateStyle1");
            style.Font.Name = "Times New Roman";
            style.Font.Color = Color.Navy;

            style = template.Styles.Add(StyleType.Paragraph, "TemplateStyle2");
            style.Font.Name = "Arial";
            style.Font.Color = Color.DeepSkyBlue;

            style = template.Styles.Add(StyleType.Paragraph, "TemplateStyle3");
            style.Font.Name = "Courier New";
            style.Font.Color = Color.RoyalBlue;

            Assert.AreEqual(7, template.Styles.Count);

            // Create a document which we will copy the styles to.
            Document target = new Document();

            // Create a style with the same name as a style from the template document and add it to the target document.
            style = target.Styles.Add(StyleType.Paragraph, "TemplateStyle3");
            style.Font.Name = "Calibri";
            style.Font.Color = Color.Orange;

            Assert.AreEqual(5, target.Styles.Count);

            // There are two ways of calling the method to copy all the styles from one document to another.
            // 1 -  Passing the template document object:
            target.CopyStylesFromTemplate(template);

            // Copying styles adds all styles from the template document to the target
            // and overwrites existing styles with the same name.
            Assert.AreEqual(7, target.Styles.Count);

            Assert.AreEqual("Courier New", target.Styles["TemplateStyle3"].Font.Name);
            Assert.AreEqual(Color.RoyalBlue.ToArgb(), target.Styles["TemplateStyle3"].Font.Color.ToArgb());

            // 2 -  Passing the local system filename of a template document:
            target.CopyStylesFromTemplate(MyDir + "Rendering.docx");

            Assert.AreEqual(21, target.Styles.Count);
            //ExEnd
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
            //ExSummary:Shows how to access a document's VBA project information.
            Document doc = new Document(MyDir + "VBA project.docm");

            // A VBA project contains a collection of VBA modules.
            VbaProject vbaProject = doc.VbaProject;
            Assert.True(vbaProject.IsSigned); //ExSkip
            Console.WriteLine(vbaProject.IsSigned
                ? $"Project name: {vbaProject.Name} signed; Project code page: {vbaProject.CodePage}; Modules count: {vbaProject.Modules.Count()}\n"
                : $"Project name: {vbaProject.Name} not signed; Project code page: {vbaProject.CodePage}; Modules count: {vbaProject.Modules.Count()}\n");

            VbaModuleCollection vbaModules = doc.VbaProject.Modules; 

            Assert.AreEqual(vbaModules.Count(), 3);

            foreach (VbaModule module in vbaModules)
                Console.WriteLine($"Module name: {module.Name};\nModule code:\n{module.SourceCode}\n");

            // Set new source code for VBA module. You can access VBA modules in the collection either by index or by name.
            vbaModules[0].SourceCode = "Your VBA code...";
            vbaModules["Module1"].SourceCode = "Your VBA code...";

            // Remove a module from the collection.
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
            //ExSummary:Shows how to access output parameters of a document's save operation.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // After we save a document, we can access the Internet Media Type (MIME type) of the newly created output document.
            SaveOutputParameters parameters = doc.Save(ArtifactsDir + "Document.SaveOutputParameters.doc");

            Assert.AreEqual("application/msword", parameters.ContentType);

            // This property changes depending on the save format.
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

            // This node serves as a reference to an external document, and its contents cannot be accessed.
            SubDocument subDocument = (SubDocument)subDocuments[0];

            Assert.False(subDocument.IsComposite);
            //ExEnd
        }

        [Test]
        public void CreateWebExtension()
        {
            //ExStart
            //ExFor:BaseWebExtensionCollection`1.Add(`0)
            //ExFor:BaseWebExtensionCollection`1.Clear
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
            //ExSummary:Shows how to add a web extension to a document.
            Document doc = new Document();

            // Create task pane with "MyScript" add-in, which will be used by the document,
            // then set its default location.
            TaskPane myScriptTaskPane = new TaskPane();
            doc.WebExtensionTaskPanes.Add(myScriptTaskPane);
            myScriptTaskPane.DockState = TaskPaneDockState.Right;
            myScriptTaskPane.IsVisible = true;
            myScriptTaskPane.Width = 300;
            myScriptTaskPane.IsLocked = true;

            // If there are multiple task panes in the same docking location, we can set this index to arrange them.
            myScriptTaskPane.Row = 1;

            // Create an add-in called "MyScript Math Sample", which the task pane will display within.
            WebExtension webExtension = myScriptTaskPane.WebExtension;

            // Set application store reference parameters for our add-in, such as the ID.
            webExtension.Reference.Id = "WA104380646";
            webExtension.Reference.Version = "1.0.0.0";
            webExtension.Reference.StoreType = WebExtensionStoreType.OMEX;
            webExtension.Reference.Store = CultureInfo.CurrentCulture.Name;
            webExtension.Properties.Add(new WebExtensionProperty("MyScript", "MyScript Math Sample"));
            webExtension.Bindings.Add(new WebExtensionBinding("MyScript", WebExtensionBindingType.Text, "104380646"));

            // Allow the user to interact with the add-in.
            webExtension.IsFrozen = false;

            // We can access the web extension in Microsoft Word via Developer -> Add-ins.
            doc.Save(ArtifactsDir + "Document.WebExtension.docx");

            // Remove all web extension task panes at once like this.
            doc.WebExtensionTaskPanes.Clear();

            Assert.AreEqual(0, doc.WebExtensionTaskPanes.Count);
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
            //ExFor:BaseWebExtensionCollection`1.GetEnumerator
            //ExFor:BaseWebExtensionCollection`1.Remove(Int32)
            //ExFor:BaseWebExtensionCollection`1.Count
            //ExFor:BaseWebExtensionCollection`1.Item(Int32)
            //ExSummary:Shows how to work with a document's collection of web extensions.
            Document doc = new Document(MyDir + "Web extension.docx");

            Assert.AreEqual(1, doc.WebExtensionTaskPanes.Count);

            // Print all properties of the document's web extension.
            WebExtensionPropertyCollection webExtensionPropertyCollection = doc.WebExtensionTaskPanes[0].WebExtension.Properties;
            using (IEnumerator<WebExtensionProperty> enumerator = webExtensionPropertyCollection.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    WebExtensionProperty webExtensionProperty = enumerator.Current;
                    Console.WriteLine($"Binding name: {webExtensionProperty.Name}; Binding value: {webExtensionProperty.Value}");
                }
            }

            // Remove the web extension.
            doc.WebExtensionTaskPanes.Remove(0);

            Assert.AreEqual(0, doc.WebExtensionTaskPanes.Count);
            //ExEnd
		}

		[Test]
        public void EpubCover()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // When saving to .epub, some Microsoft Word document properties convert to .epub metadata.
            doc.BuiltInDocumentProperties.Author = "John Doe";
            doc.BuiltInDocumentProperties.Title = "My Book Title";

            // The thumbnail we specify here can become the cover image.
            byte[] image = File.ReadAllBytes(ImageDir + "Transparent background logo.png");
            doc.BuiltInDocumentProperties.Thumbnail = image;

            doc.Save(ArtifactsDir + "Document.EpubCover.epub");
        }
        
        [Test]
        public void TextWatermark()
        {
            //ExStart
            //ExFor:Watermark.SetText(String)
            //ExFor:Watermark.SetText(String, TextWatermarkOptions)
            //ExFor:Watermark.Remove
            //ExFor:TextWatermarkOptions.FontFamily
            //ExFor:TextWatermarkOptions.FontSize
            //ExFor:TextWatermarkOptions.Color
            //ExFor:TextWatermarkOptions.Layout
            //ExFor:TextWatermarkOptions.IsSemitrasparent
            //ExFor:WatermarkLayout
            //ExFor:WatermarkType
            //ExSummary:Shows how to create a text watermark.
            Document doc = new Document();

            // Add a plain text watermark.
            doc.Watermark.SetText("Aspose Watermark");
            
            // If we wish to edit the text formatting using it as a watermark,
            // we can do so by passing a TextWatermarkOptions object when creating the watermark.
            TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
            textWatermarkOptions.FontFamily = "Arial";
            textWatermarkOptions.FontSize = 36;
            textWatermarkOptions.Color = Color.Black;
            textWatermarkOptions.Layout = WatermarkLayout.Diagonal;
            textWatermarkOptions.IsSemitrasparent = false;

            doc.Watermark.SetText("Aspose Watermark", textWatermarkOptions);

            doc.Save(ArtifactsDir + "Document.TextWatermark.docx");

            // We can remove a watermark from a document like this.
            if (doc.Watermark.Type == WatermarkType.Text)
                doc.Watermark.Remove();
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.TextWatermark.docx");

            Assert.AreEqual(WatermarkType.Text, doc.Watermark.Type);
        }

        [Test]
        public void ImageWatermark()
        {
            //ExStart
            //ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
            //ExFor:ImageWatermarkOptions.Scale
            //ExFor:ImageWatermarkOptions.IsWashout
            //ExSummary:Shows how to create a watermark from an image in the local file system.
            Document doc = new Document();

            // Modify the image watermark's appearance with an ImageWatermarkOptions object,
            // then pass it while creating a watermark from an image file.
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

            doc.Save(ArtifactsDir + "Document.ImageWatermark.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.ImageWatermark.docx");

            Assert.AreEqual(WatermarkType.Image, doc.Watermark.Type);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void SpellingAndGrammarErrors(bool showErrors)
        {
            //ExStart
            //ExFor:Document.ShowGrammaticalErrors
            //ExFor:Document.ShowSpellingErrors
            //ExSummary:Shows how to show/hide errors in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert two sentences with mistakes that would be picked up
            // by the spelling and grammar checkers in Microsoft Word.
            builder.Writeln("There is a speling error in this sentence.");
            builder.Writeln("Their is a grammatical error in this sentence.");

            // If these options are enabled, then spelling errors will be underlined
            // in the output document by a jagged red line, and a double blue line will highlight grammatical mistakes.
            doc.ShowGrammaticalErrors = showErrors;
            doc.ShowSpellingErrors = showErrors;
            
            doc.Save(ArtifactsDir + "Document.SpellingAndGrammarErrors.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.SpellingAndGrammarErrors.docx");

            Assert.AreEqual(showErrors, doc.ShowGrammaticalErrors);
            Assert.AreEqual(showErrors, doc.ShowSpellingErrors);
        }

        [TestCase(Granularity.CharLevel)]
        [TestCase(Granularity.WordLevel)]
        public void GranularityCompareOption(Granularity granularity)
        {
            //ExStart
            //ExFor:CompareOptions.Granularity
            //ExFor:Granularity
            //ExSummary:Shows to specify a granularity while comparing documents.
            Document docA = new Document();
            DocumentBuilder builderA = new DocumentBuilder(docA);
            builderA.Writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");

            Document docB = new Document();
            DocumentBuilder builderB = new DocumentBuilder(docB);
            builderB.Writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");
 
            // Specify whether changes are tracking
            // by character ('Granularity.CharLevel'), or by word ('Granularity.WordLevel').
            CompareOptions compareOptions = new CompareOptions();
            compareOptions.Granularity = granularity;
 
            docA.Compare(docB, "author", DateTime.Now, compareOptions);

            // The first document's collection of revision groups contains all the differences between documents.
            RevisionGroupCollection groups = docA.Revisions.Groups;
            Assert.AreEqual(5, groups.Count);
            //ExEnd

            if (granularity == Granularity.CharLevel)
            {
                Assert.AreEqual(RevisionType.Deletion, groups[0].RevisionType);
                Assert.AreEqual("Alpha ", groups[0].Text);

                Assert.AreEqual(RevisionType.Deletion, groups[1].RevisionType);
                Assert.AreEqual(",", groups[1].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[2].RevisionType);
                Assert.AreEqual("s", groups[2].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[3].RevisionType);
                Assert.AreEqual("- \"", groups[3].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[4].RevisionType);
                Assert.AreEqual("\"", groups[4].Text);
            }
            else
            {
                Assert.AreEqual(RevisionType.Deletion, groups[0].RevisionType);
                Assert.AreEqual("Alpha Lorem ", groups[0].Text);

                Assert.AreEqual(RevisionType.Deletion, groups[1].RevisionType);
                Assert.AreEqual(",", groups[1].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[2].RevisionType);
                Assert.AreEqual("Lorems ", groups[2].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[3].RevisionType);
                Assert.AreEqual("- \"", groups[3].Text);

                Assert.AreEqual(RevisionType.Insertion, groups[4].RevisionType);
                Assert.AreEqual("\"", groups[4].Text);   
            }
        }

        [Test]
        public void IgnorePrinterMetrics()
        {
            //ExStart
            //ExFor:LayoutOptions.IgnorePrinterMetrics
            //ExSummary:Shows how to ignore 'Use printer metrics to lay out document' option.
            Document doc = new Document(MyDir + "Rendering.docx");
    
            doc.LayoutOptions.IgnorePrinterMetrics = false;

            doc.Save(ArtifactsDir + "Document.IgnorePrinterMetrics.docx");
            //ExEnd
        }

        [Test]
        public void ExtractPages()
        {
            //ExStart
            //ExFor:Document.ExtractPages
            //ExSummary:Shows how to get specified range of pages from the document.
            Document doc = new Document(MyDir + "Layout entities.docx");

            doc = doc.ExtractPages(0, 2);
    
            doc.Save(ArtifactsDir + "Document.ExtractPages.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.ExtractPages.docx");
            Assert.AreEqual(doc.PageCount, 2);
        }

        [TestCase(true)]
        [TestCase(false)]
        public void SpellingOrGrammar(bool checkSpellingGrammar)
        {
            //ExStart
            //ExFor:Document.SpellingChecked
            //ExFor:Document.GrammarChecked
            //ExSummary:Shows how to set spelling or grammar verifying.
            Document doc = new Document();

            // The string with spelling errors.
            doc.FirstSection.Body.FirstParagraph.Runs.Add(new Run(doc, "The speeling in this documentz is all broked."));

            // Spelling/Grammar check start if we set properties to false. 
            // We can see all errors in Microsoft Word via Review -> Spelling & Grammar.
            // Note that Microsoft Word does not start grammar/spell check automatically for DOC and RTF document format.
            doc.SpellingChecked = checkSpellingGrammar;
            doc.GrammarChecked = checkSpellingGrammar;

            doc.Save(ArtifactsDir + "Document.SpellingOrGrammar.docx");
            //ExEnd
        }

        [Test]
        public void AllowEmbeddingPostScriptFonts()
        {
            //ExStart
            //ExFor:SaveOptions.AllowEmbeddingPostScriptFonts
            //ExSummary:Shows how to save the document with PostScript font.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "PostScriptFont";
            builder.Writeln("Some text with PostScript font.");

            // Load the font with PostScript to use in the document.
            MemoryFontSource otf = new MemoryFontSource(File.ReadAllBytes(FontsDir + "AllegroOpen.otf"));
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { otf });

            // Embed TrueType fonts.
            doc.FontInfos.EmbedTrueTypeFonts = true;

            // Allow embedding PostScript fonts while embedding TrueType fonts.
            // Microsoft Word does not embed PostScript fonts, but can open documents with embedded fonts of this type.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
            saveOptions.AllowEmbeddingPostScriptFonts = true;

            doc.Save(ArtifactsDir + "Document.AllowEmbeddingPostScriptFonts.docx", saveOptions);
            //ExEnd
        }
    }
}