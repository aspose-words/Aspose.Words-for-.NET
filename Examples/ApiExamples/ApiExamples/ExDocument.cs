// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
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
using MemoryFontSource = Aspose.Words.Fonts.MemoryFontSource;
using LoadOptions = Aspose.Words.Loading.LoadOptions;
using Aspose.Words.Settings;
using Aspose.Pdf.Text;
using Aspose.Words.Shaping.HarfBuzz;
using System.Net.Http;
#if NET5_0_OR_GREATER || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDocument : ApiExampleBase
    {
        [Test]
        public void CreateSimpleDocument()
        {
            //ExStart:CreateSimpleDocument
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:DocumentBase.Document
            //ExFor:Document.#ctor()
            //ExSummary:Shows how to create simple document.
            Document doc = new Document();

            // New Document objects by default come with the minimal set of nodes
            // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
            doc.AppendChild(new Section(doc))
                .AppendChild(new Body(doc))
                .AppendChild(new Paragraph(doc))
                .AppendChild(new Run(doc, "Hello world!"));
            //ExEnd:CreateSimpleDocument
        }

        [Test]
        public void Constructor()
        {
            //ExStart
            //ExFor:Document.#ctor()
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
            Assert.That(doc.FirstSection.Body.FirstParagraph.GetText().Trim(), Is.EqualTo("Hello World!"));

            // Some operations that need to occur during loading, such as using a password to decrypt a document,
            // can be done by passing a LoadOptions object when loading the document.
            doc = new Document(MyDir + "Encrypted.docx", new LoadOptions("docPassword"));

            Assert.That(doc.FirstSection.Body.FirstParagraph.GetText().Trim(), Is.EqualTo("Test encrypted document."));
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

                Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello World!\r\rHello Word!\r\r\rHello World!"));
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
            const string url = "https://filesamples.com/samples/document/docx/sample3.docx";

            // Download the document into a byte array, then load that array into a document using a memory stream.
            using (HttpClient httpClient = new HttpClient())
            {
                HttpResponseMessage response = httpClient.GetAsync(url).Result;
                byte[] dataBytes = response.Content.ReadAsByteArrayAsync().Result;

                using (MemoryStream byteStream = new MemoryStream(dataBytes))
                {
                    Document doc = new Document(byteStream);

                    // At this stage, we can read and edit the document's contents and then save it to the local file system.
                    Assert.That(doc.FirstSection.Body.Paragraphs[3].GetText().Trim(), Is.EqualTo("There are eight section headings in this document. At the beginning, \"Sample Document\" is a level 1 heading. " +
                                  "The main section headings, such as \"Headings\" and \"Lists\" are level 2 headings. " +
                                    "The Tables section contains two sub-headings, \"Simple Table\" and \"Complex Table,\" which are both level 3 headings."));

                    doc.Save(ArtifactsDir + "Document.LoadFromWeb.docx");
                }
            }
            //ExEnd
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

#if NET461_OR_GREATER || JAVA
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Bmp);

                stream.Position = 0;

                // Read the stream back into an image.
                using (Image image = Image.FromStream(stream))
                {
                    Assert.That(image.RawFormat, Is.EqualTo(ImageFormat.Bmp));
                    Assert.That(image.Width, Is.EqualTo(816));
                    Assert.That(image.Height, Is.EqualTo(1056));
                }
            }
#elif NET6_0_OR_GREATER
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Bmp);

                stream.Position = 0;

                SKCodec codec = SKCodec.Create(stream);
                Assert.That(SKEncodedImageFormat.Bmp, Is.EqualTo(codec.EncodedFormat));

                stream.Position = 0;

                using (SKBitmap image = SKBitmap.Decode(stream))
                {
                    Assert.That(816, Is.EqualTo(image.Width));
                    Assert.That(1056, Is.EqualTo(image.Height));
                }
            }
#endif
            //ExEnd
        }

        [Test, Category("SkipMono")]
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

        [Test]
        public void DetectMobiDocumentFormat()
        {
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Document.mobi");
            Assert.That(LoadFormat.Mobi, Is.EqualTo(info.LoadFormat));
        }

        [Test]
        public void DetectPdfDocumentFormat()
        {
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Pdf Document.pdf");
            Assert.That(LoadFormat.Pdf, Is.EqualTo(info.LoadFormat));
        }

        [Test]
        public void OpenPdfDocument()
        {
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            Assert.That(doc.Range.Text, Is.EqualTo("Heading 1\rHeading 1.1.1.1 Heading 1.1.1.2\rHeading 1.1.1.1.1.1.1.1.1 Heading 1.1.1.1.1.1.1.1.2\u000c"));
        }

        [Test]
        public void OpenProtectedPdfDocument()
        {
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EncryptionDetails = new PdfEncryptionDetails("Aspose", null);

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
            //ExFor:ShapeBase.IsImage
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

                Assert.That(shape.IsImage, Is.True);
                Assert.That(shape.ImageData.ImageBytes, Is.Not.Null);
                Assert.That(ConvertUtil.PointToPixel(shape.Width), Is.EqualTo(32.0).Within(0.01));
                Assert.That(ConvertUtil.PointToPixel(shape.Height), Is.EqualTo(32.0).Within(0.01));
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
            //ExSummary:Shows how save a web page as a .docx file.
            const string url = "https://products.aspose.com/words/";

            using (HttpClient client = new HttpClient())
            {
                byte[] bytes = client.GetByteArrayAsync(url).GetAwaiter().GetResult();

                using (MemoryStream stream = new MemoryStream(bytes))
                {
                    // The URL is used again as a baseUri to ensure that any relative image paths are retrieved correctly.
                    LoadOptions options = new LoadOptions(LoadFormat.Html, "", url);

                    // Load the HTML document from stream and pass the LoadOptions object.
                    Document doc = new Document(stream, options);

                    // At this stage, we can read and edit the document's contents and then save it to the local file system.
                    Assert.That(doc.GetText().Contains("HYPERLINK \"https://products.aspose.com/words/net/\" \\o \"Aspose.Words\""), Is.True); //ExSkip

                    doc.Save(ArtifactsDir + "Document.InsertHtmlFromWebPage.docx");
                }
            }
            //ExEnd
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
            Assert.That(doc.GetText().Trim(), Is.EqualTo("Test encrypted document.")); //ExSkip

            // 2 -  Load the document from a stream:
            using (Stream stream = File.OpenRead(MyDir + "Encrypted.docx"))
            {
                doc = new Document(stream, options);
                Assert.That(doc.GetText().Trim(), Is.EqualTo("Test encrypted document.")); //ExSkip
            }
            //ExEnd
        }

        [Test]
        public void NotSupportedWarning()
        {
            //ExStart
            //ExFor:WarningInfoCollection.Count
            //ExFor:WarningInfoCollection.Item(Int32)
            //ExSummary:Shows how to get warnings about unsupported formats.
            WarningInfoCollection warnings = new WarningInfoCollection();
            Document doc = new Document(MyDir + "FB2 document.fb2", new LoadOptions { WarningCallback = warnings });

            Assert.That(warnings[0].Description, Is.EqualTo("The original file load format is FB2, which is not supported by Aspose.Words. The file is loaded as an XML document."));
            Assert.That(warnings.Count, Is.EqualTo(1));
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
                Assert.That(new Document(dstStream).GetText().Trim(), Is.EqualTo("Hello World!\r\rHello Word!\r\r\rHello World!"));
            }
            //ExEnd
        }

        //ExStart
        //ExFor:Range.Fields
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
                    Aspose.Words.Font font = ((Run)args.Node).Font;
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
            Assert.That(Regex.Matches(log, "insertion").Count, Is.EqualTo(10));
            Assert.That(Regex.Matches(log, "removal").Count, Is.EqualTo(5));
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
            Assert.That(dstDoc.Sections.Count, Is.EqualTo(2)); //ExSkip

            dstDoc.Save(ArtifactsDir + "Document.AppendDocument.docx");
            //ExEnd

            string outDocText = new Document(ArtifactsDir + "Document.AppendDocument.docx").GetText();

            Assert.That(outDocText.StartsWith(dstDoc.GetText()), Is.True);
            Assert.That(outDocText.EndsWith(srcDoc.GetText()), Is.True);
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

                Assert.Throws<FileNotFoundException>(() => new Document("C:\\DetailsList.doc"));

                // Append the source document at the end of the destination document.
                doc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

                // Automation required you to insert a new section break at this point, however, in Aspose.Words we
                // do not need to do anything here as the appended document is imported as separate sections already

                // Unlink all headers/footers in this section from the previous section headers/footers
                // if this is the second document or above being appended.
                if (i > 1)
                    Assert.Throws<NullReferenceException>(() => doc.Sections[i].HeadersFooters.LinkToPrevious(false));
            }
        }

        [TestCase(true)]
        [TestCase(false)]
        public void ImportList(bool isKeepSourceNumbering)
        {
            //ExStart
            //ExFor:ImportFormatOptions.KeepSourceNumbering
            //ExSummary:Shows how to import a document with numbered lists.
            Document srcDoc = new Document(MyDir + "List source.docx");
            Document dstDoc = new Document(MyDir + "List destination.docx");

            Assert.That(dstDoc.Lists.Count, Is.EqualTo(4));

            ImportFormatOptions options = new ImportFormatOptions();

            // If there is a clash of list styles, apply the list format of the source document.
            // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
            // Set the "KeepSourceNumbering" property to "true" import all clashing
            // list style numbering with the same appearance that it had in the source document.
            options.KeepSourceNumbering = isKeepSourceNumbering;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting, options);
            dstDoc.UpdateListLabels();

            Assert.That(dstDoc.Lists.Count, Is.EqualTo(isKeepSourceNumbering ? 5 : 4));
            //ExEnd
        }

        [Test]
        public void KeepSourceNumberingSameListIds()
        {
            //ExStart
            //ExFor:ImportFormatOptions.KeepSourceNumbering
            //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
            //ExSummary:Shows how resolve a clash when importing documents that have lists with the same list definition identifier.
            Document srcDoc = new Document(MyDir + "List with the same definition identifier - source.docx");
            Document dstDoc = new Document(MyDir + "List with the same definition identifier - destination.docx");

            // Set the "KeepSourceNumbering" property to "true" to apply a different list definition ID
            // to identical styles as Aspose.Words imports them into destination documents.
            ImportFormatOptions importFormatOptions = new ImportFormatOptions { KeepSourceNumbering = true };

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles, importFormatOptions);
            dstDoc.UpdateListLabels();
            //ExEnd

            string paraText = dstDoc.Sections[1].Body.LastParagraph.GetText();

            Assert.That(paraText.StartsWith("13->13"), Is.True, paraText);
            Assert.That(dstDoc.Sections[1].Body.LastParagraph.ListLabel.LabelString, Is.EqualTo("1."));
        }

        [Test]
        public void MergePastedLists()
        {
            //ExStart
            //ExFor:ImportFormatOptions.MergePastedLists
            //ExSummary:Shows how to merge lists from a documents.
            Document srcDoc = new Document(MyDir + "List item.docx");
            Document dstDoc = new Document(MyDir + "List destination.docx");

            ImportFormatOptions options = new ImportFormatOptions { MergePastedLists = true };

            // Set the "MergePastedLists" property to "true" pasted lists will be merged with surrounding lists.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);

            dstDoc.Save(ArtifactsDir + "Document.MergePastedLists.docx");
            //ExEnd
        }

        [Test]
        public void ForceCopyStyles()
        {
            //ExStart
            //ExFor:ImportFormatOptions.ForceCopyStyles
            //ExSummary:Shows how to copy source styles with unique names forcibly.
            // Both documents contain MyStyle1 and MyStyle2, MyStyle3 exists only in a source document.
            Document srcDoc = new Document(MyDir + "Styles source.docx");
            Document dstDoc = new Document(MyDir + "Styles destination.docx");

            ImportFormatOptions options = new ImportFormatOptions { ForceCopyStyles = true };
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting, options);

            ParagraphCollection paras = dstDoc.Sections[1].Body.Paragraphs;

            Assert.That("MyStyle1_0", Is.EqualTo(paras[0].ParagraphFormat.Style.Name));
            Assert.That("MyStyle2_0", Is.EqualTo(paras[1].ParagraphFormat.Style.Name));
            Assert.That("MyStyle3", Is.EqualTo(paras[2].ParagraphFormat.Style.Name));
            //ExEnd
        }

        [Test]
        public void AdjustSentenceAndWordSpacing()
        {
            //ExStart
            //ExFor:ImportFormatOptions.AdjustSentenceAndWordSpacing
            //ExSummary:Shows how to adjust sentence and word spacing automatically.
            Document srcDoc = new Document();
            Document dstDoc = new Document();

            DocumentBuilder builder = new DocumentBuilder(srcDoc);
            builder.Write("Dolor sit amet.");

            builder = new DocumentBuilder(dstDoc);
            builder.Write("Lorem ipsum.");

            ImportFormatOptions options = new ImportFormatOptions() { AdjustSentenceAndWordSpacing = true };
            builder.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);

            Assert.That(dstDoc.FirstSection.Body.FirstParagraph.GetText().Trim(), Is.EqualTo("Lorem ipsum. Dolor sit amet."));
            //ExEnd
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

            Assert.That(doc.DigitalSignatures.Count, Is.EqualTo(1));

            DigitalSignature digitalSig = doc.DigitalSignatures[0];

            Assert.That(digitalSig.IsValid, Is.True);
            Assert.That(digitalSig.Comments, Is.EqualTo("Test Sign"));
            Assert.That(digitalSig.SignatureType.ToString(), Is.EqualTo("XmlDsig"));
            Assert.That(digitalSig.CertificateHolder.Certificate.Subject.Contains("Aspose Pty Ltd"), Is.True);
            Assert.That(digitalSig.CertificateHolder.Certificate.IssuerName.Name != null &&
                        digitalSig.CertificateHolder.Certificate.IssuerName.Name.Contains("VeriSign"), Is.True);
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
            Assert.That(FileFormatUtil.DetectFileFormat(MyDir + "Document.docx").HasDigitalSignature, Is.False);

            // Create a CertificateHolder object from a PKCS12 file, which we will use to sign the document.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);

            // There are two ways of saving a signed copy of a document to the local file system:
            // 1 - Designate a document by a local system filename and save a signed copy at a location specified by another filename.
            SignOptions signOptions = new SignOptions { SignTime = DateTime.Now };
            DigitalSignatureUtil.Sign(MyDir + "Document.docx", ArtifactsDir + "Document.DigitalSignature.docx",
                certificateHolder, signOptions);

            Assert.That(FileFormatUtil.DetectFileFormat(ArtifactsDir + "Document.DigitalSignature.docx").HasDigitalSignature, Is.True);

            // 2 - Take a document from a stream and save a signed copy to another stream.
            using (FileStream inDoc = new FileStream(MyDir + "Document.docx", FileMode.Open))
            {
                using (FileStream outDoc = new FileStream(ArtifactsDir + "Document.DigitalSignature.docx", FileMode.Create))
                {
                    DigitalSignatureUtil.Sign(inDoc, outDoc, certificateHolder);
                }
            }

            Assert.That(FileFormatUtil.DetectFileFormat(ArtifactsDir + "Document.DigitalSignature.docx").HasDigitalSignature, Is.True);

            // Please verify that all of the document's digital signatures are valid and check their details.
            Document signedDoc = new Document(ArtifactsDir + "Document.DigitalSignature.docx");
            DigitalSignatureCollection digitalSignatureCollection = signedDoc.DigitalSignatures;

            Assert.That(digitalSignatureCollection.IsValid, Is.True);
            Assert.That(digitalSignatureCollection.Count, Is.EqualTo(1));
            Assert.That(digitalSignatureCollection[0].SignatureType, Is.EqualTo(DigitalSignatureType.XmlDsig));
            Assert.That(signedDoc.DigitalSignatures[0].IssuerName, Is.EqualTo("CN=Morzal.Me"));
            Assert.That(signedDoc.DigitalSignatures[0].SubjectName, Is.EqualTo("CN=Morzal.Me"));
            //ExEnd
        }

        [Test]
        public void SignatureValue()
        {
            //ExStart
            //ExFor:DigitalSignature.SignatureValue
            //ExSummary:Shows how to get a digital signature value from a digitally signed document.
            Document doc = new Document(MyDir + "Digitally signed.docx");

            foreach (DigitalSignature digitalSignature in doc.DigitalSignatures)
            {
                string signatureValue = Convert.ToBase64String(digitalSignature.SignatureValue);
                Assert.That(signatureValue, Is.EqualTo("K1cVLLg2kbJRAzT5WK+m++G8eEO+l7S+5ENdjMxxTXkFzGUfvwxREuJdSFj9AbD" +
                    "MhnGvDURv9KEhC25DDF1al8NRVR71TF3CjHVZXpYu7edQS5/yLw/k5CiFZzCp1+MmhOdYPcVO+Fm" +
                    "+9fKr2iNLeyYB+fgEeZHfTqTFM2WwAqo="));
            }
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
            Assert.That(dstDoc.Styles.Count, Is.EqualTo(5)); //ExSkip
            Assert.That(dstDoc.Sections.Count, Is.EqualTo(1)); //ExSkip

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

            Assert.That(dstDoc.Styles.Count, Is.EqualTo(7));
            Assert.That(dstDoc.Sections.Count, Is.EqualTo(10));
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
            Assert.That(doc.GetChildNodes(NodeType.Run, true).Count, Is.EqualTo(317));

            // Combine such runs with this method and verify the number of run joins that will take place.
            Assert.That(doc.JoinRunsWithSameFormatting(), Is.EqualTo(121));

            // The number of joins and the number of runs we have after the join
            // should add up the number of runs we had initially.
            Assert.That(doc.GetChildNodes(NodeType.Run, true).Count, Is.EqualTo(196));
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
            Assert.That(doc.DefaultTabStop, Is.EqualTo(72));
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

            Assert.That(clone.FirstSection.Body.FirstParagraph.Runs[0].Text, Is.EqualTo(doc.FirstSection.Body.FirstParagraph.Runs[0].GetText()));
            Assert.That(clone.FirstSection.Body.FirstParagraph.Runs[0].GetHashCode(), Is.Not.EqualTo(doc.FirstSection.Body.FirstParagraph.Runs[0].GetHashCode()));
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
            Assert.That(doc.GetText().Trim(), Is.EqualTo("\u0013MERGEFIELD Field\u0014«Field»\u0015"));

            // ToString will give us the document's appearance if saved to a passed save format.
            Assert.That(doc.ToString(SaveFormat.Text).Trim(), Is.EqualTo("«Field»"));
            //ExEnd
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

            Assert.That(doc.ProtectionType, Is.EqualTo(ProtectionType.ReadOnly));

            // If we open this document with Microsoft Word intending to edit it,
            // we will need to apply the password to get through the protection.
            doc.Save(ArtifactsDir + "Document.Protect.docx");

            // Note that the protection only applies to Microsoft Word users opening our document.
            // We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
            Document protectedDoc = new Document(ArtifactsDir + "Document.Protect.docx");

            Assert.That(protectedDoc.ProtectionType, Is.EqualTo(ProtectionType.ReadOnly));

            DocumentBuilder builder = new DocumentBuilder(protectedDoc);
            builder.Writeln("Text added to a protected document.");
            Assert.That(protectedDoc.Range.Text.Trim(), Is.EqualTo("Text added to a protected document.")); //ExSkip

            // There are two ways of removing protection from a document.
            // 1 - With no password:
            doc.Unprotect();

            Assert.That(doc.ProtectionType, Is.EqualTo(ProtectionType.NoProtection));

            doc.Protect(ProtectionType.ReadOnly, "NewPassword");

            Assert.That(doc.ProtectionType, Is.EqualTo(ProtectionType.ReadOnly));

            doc.Unprotect("WrongPassword");

            Assert.That(doc.ProtectionType, Is.EqualTo(ProtectionType.ReadOnly));

            // 2 - With the correct password:
            doc.Unprotect("NewPassword");

            Assert.That(doc.ProtectionType, Is.EqualTo(ProtectionType.NoProtection));
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

            Assert.That(nodes[0].NodeType, Is.EqualTo(NodeType.Section));
            Assert.That(nodes[0].ParentNode, Is.EqualTo(doc));

            Assert.That(nodes[1].NodeType, Is.EqualTo(NodeType.Body));
            Assert.That(nodes[1].ParentNode, Is.EqualTo(nodes[0]));

            Assert.That(nodes[2].NodeType, Is.EqualTo(NodeType.Paragraph));
            Assert.That(nodes[2].ParentNode, Is.EqualTo(nodes[1]));

            // This is the minimal set of nodes that we need to be able to edit the document.
            // We will no longer be able to edit the document if we remove any of them.
            doc.RemoveAllChildren();

            Assert.That(doc.GetChildNodes(NodeType.Any, true).Count, Is.EqualTo(0));

            // Call this method to make sure that the document has at least those three nodes so we can edit it again.
            doc.EnsureMinimum();

            Assert.That(nodes[0].NodeType, Is.EqualTo(NodeType.Section));
            Assert.That(nodes[1].NodeType, Is.EqualTo(NodeType.Body));
            Assert.That(nodes[2].NodeType, Is.EqualTo(NodeType.Paragraph));

            ((Paragraph)nodes[2]).Runs.Add(new Run(doc, "Hello world!"));
            //ExEnd

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello world!"));
        }

        [Test]
        public void RemoveMacrosFromDocument()
        {
            //ExStart
            //ExFor:Document.RemoveMacros
            //ExSummary:Shows how to remove all macros from a document.
            Document doc = new Document(MyDir + "Macro.docm");

            Assert.That(doc.HasMacros, Is.True);
            Assert.That(doc.VbaProject.Name, Is.EqualTo("Project"));

            // Remove the document's VBA project, along with all its macros.
            doc.RemoveMacros();

            Assert.That(doc.HasMacros, Is.False);
            Assert.That(doc.VbaProject, Is.Null);
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
            Assert.That(doc.PageCount, Is.EqualTo(3));

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
            Assert.That(doc.BuiltInDocumentProperties.Characters, Is.EqualTo(0));
            Assert.That(doc.BuiltInDocumentProperties.Words, Is.EqualTo(0));
            Assert.That(doc.BuiltInDocumentProperties.Paragraphs, Is.EqualTo(1));
            Assert.That(doc.BuiltInDocumentProperties.Lines, Is.EqualTo(1));

            // To get accurate values for three of these properties, we will need to update them manually.
            doc.UpdateWordCount();

            Assert.That(doc.BuiltInDocumentProperties.Characters, Is.EqualTo(196));
            Assert.That(doc.BuiltInDocumentProperties.Words, Is.EqualTo(36));
            Assert.That(doc.BuiltInDocumentProperties.Paragraphs, Is.EqualTo(2));

            // For the line count, we will need to call a specific overload of the updating method.
            Assert.That(doc.BuiltInDocumentProperties.Lines, Is.EqualTo(1));

            doc.UpdateWordCount(true);

            Assert.That(doc.BuiltInDocumentProperties.Lines, Is.EqualTo(4));
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
        public void GetOriginalFileInfo()
        {
            //ExStart
            //ExFor:Document.OriginalFileName
            //ExFor:Document.OriginalLoadFormat
            //ExSummary:Shows how to retrieve details of a document's load operation.
            Document doc = new Document(MyDir + "Document.docx");

            Assert.That(doc.OriginalFileName, Is.EqualTo(MyDir + "Document.docx"));
            Assert.That(doc.OriginalLoadFormat, Is.EqualTo(LoadFormat.Docx));
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
            Assert.That(doc.FootnoteOptions.Columns, Is.EqualTo(0)); //ExSkip

            doc.FootnoteOptions.Columns = 2;
            doc.Save(ArtifactsDir + "Document.FootnoteColumns.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.FootnoteColumns.docx");

            Assert.That(doc.FirstSection.PageSetup.FootnoteOptions.Columns, Is.EqualTo(2));
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
            Assert.That(options.ThumbnailSize, Is.EqualTo(new Size(600, 900))); //ExSkip
            Assert.That(options.GenerateFromFirstPage, Is.True); //ExSkip
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

            Assert.That(doc.HyphenationOptions.AutoHyphenation, Is.EqualTo(true));
            Assert.That(doc.HyphenationOptions.ConsecutiveHyphenLimit, Is.EqualTo(2));
            Assert.That(doc.HyphenationOptions.HyphenationZone, Is.EqualTo(720));
            Assert.That(doc.HyphenationOptions.HyphenateCaps, Is.EqualTo(true));

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "Document.HyphenationOptions.docx",
                GoldsDir + "Document.HyphenationOptions Gold.docx"), Is.True);
        }

        [Test]
        public void HyphenationOptionsDefaultValues()
        {
            Document doc = new Document();
            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.HyphenationOptions.AutoHyphenation, Is.EqualTo(false));
            Assert.That(doc.HyphenationOptions.ConsecutiveHyphenLimit, Is.EqualTo(0));
            Assert.That(doc.HyphenationOptions.HyphenationZone, Is.EqualTo(360)); // 0.25 inch
            Assert.That(doc.HyphenationOptions.HyphenateCaps, Is.EqualTo(true));
        }

        [Test]
        public void HyphenationZoneException()
        {
            Document doc = new Document();

            Assert.Throws<ArgumentOutOfRangeException>(() => doc.HyphenationOptions.HyphenationZone = 0);
        }

        [Test]
        public void OoxmlComplianceVersion()
        {
            //ExStart
            //ExFor:Document.Compliance
            //ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
            // The compliance version varies between documents created by different versions of Microsoft Word.
            Document doc = new Document(MyDir + "Document.doc");
            Assert.That(OoxmlCompliance.Ecma376_2006, Is.EqualTo(doc.Compliance));

            doc = new Document(MyDir + "Document.docx");
            Assert.That(OoxmlCompliance.Iso29500_2008_Transitional, Is.EqualTo(doc.Compliance));
            //ExEnd
        }

        [Test]
        [Description("WORDSNET-20342")]
        public void ImageSaveOptions()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:SaveOptions.UseAntiAliasing
            //ExFor:SaveOptions.UseHighQualityRendering
            //ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
            Document doc = new Document(MyDir + "Rendering.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Size = 60;
            builder.Writeln("Some text.");

            SaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);
            Assert.That(options.UseAntiAliasing, Is.False); //ExSkip
            Assert.That(options.UseHighQualityRendering, Is.False); //ExSkip

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
            Assert.That(doc.Styles.Count, Is.EqualTo(8));

            // Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Style = doc.Styles["MyParagraphStyle1"];
            builder.Writeln("Hello world!");

            Aspose.Words.Lists.List docList = doc.Lists.Add(doc.Styles["MyListStyle1"]);
            builder.ListFormat.List = docList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            doc.Cleanup();

            Assert.That(doc.Styles.Count, Is.EqualTo(6));

            // Removing every node that a custom style is applied to marks it as "unused" again.
            // Run the Cleanup method again to remove them.
            doc.FirstSection.Body.RemoveAllChildren();
            doc.Cleanup();

            Assert.That(doc.Styles.Count, Is.EqualTo(4));
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
            Assert.That(doc.AttachedTemplate, Is.EqualTo(string.Empty));

            // Attach a template, then set the flag to apply style changes
            // within the template to styles in our document.
            doc.AttachedTemplate = MyDir + "Business brochure.dotx";
            doc.AutomaticallyUpdateStyles = true;

            doc.Save(ArtifactsDir + "Document.AutomaticallyUpdateStyles.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.AutomaticallyUpdateStyles.docx");

            Assert.That(doc.AutomaticallyUpdateStyles, Is.True);
            Assert.That(doc.AttachedTemplate, Is.EqualTo(MyDir + "Business brochure.dotx"));
            Assert.That(File.Exists(doc.AttachedTemplate), Is.True);
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

            Assert.That(doc.AttachedTemplate, Is.EqualTo(string.Empty));

            // Since there is no template document, the document had nowhere to track style changes.
            // Use a SaveOptions object to automatically set a template
            // if a document that we are saving does not have one.
            SaveOptions options = SaveOptions.CreateSaveOptions("Document.DefaultTemplate.docx");
            options.DefaultTemplate = MyDir + "Business brochure.dotx";

            doc.Save(ArtifactsDir + "Document.DefaultTemplate.docx", options);
            //ExEnd

            Assert.That(File.Exists(options.DefaultTemplate), Is.True);
        }

        [Test]
        public void UseSubstitutions()
        {
            //ExStart
            //ExFor:FindReplaceOptions.#ctor()
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

            Assert.That("Paul took money from Jason.\f", Is.EqualTo(doc.GetText()));
            //ExEnd
        }

        [Test]
        public void SetInvalidateFieldTypes()
        {
            //ExStart
            //ExFor:Document.NormalizeFieldTypes
            //ExFor:Range.NormalizeFieldTypes
            //ExSummary:Shows how to get the keep a field's type up to date with its field code.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Field field = builder.InsertField("DATE", null);

            // Aspose.Words automatically detects field types based on field codes.
            Assert.That(field.Type, Is.EqualTo(FieldType.FieldDate));

            // Manually change the raw text of the field, which determines the field code.
            Run fieldText = (Run)doc.FirstSection.Body.FirstParagraph.GetChildNodes(NodeType.Run, true)[0];
            Assert.That(fieldText.Text, Is.EqualTo("DATE")); //ExSkip
            fieldText.Text = "PAGE";

            // Changing the field code has changed this field to one of a different type,
            // but the field's type properties still display the old type.
            Assert.That(field.GetFieldCode(), Is.EqualTo("PAGE"));
            Assert.That(field.Type, Is.EqualTo(FieldType.FieldDate));
            Assert.That(field.Start.FieldType, Is.EqualTo(FieldType.FieldDate));
            Assert.That(field.Separator.FieldType, Is.EqualTo(FieldType.FieldDate));
            Assert.That(field.End.FieldType, Is.EqualTo(FieldType.FieldDate));

            // Update those properties with this method to display current value.
            doc.NormalizeFieldTypes();

            Assert.That(field.Type, Is.EqualTo(FieldType.FieldPage));
            Assert.That(field.Start.FieldType, Is.EqualTo(FieldType.FieldPage));
            Assert.That(field.Separator.FieldType, Is.EqualTo(FieldType.FieldPage));
            Assert.That(field.End.FieldType, Is.EqualTo(FieldType.FieldPage));
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void LayoutOptionsHiddenText(bool showHiddenText)
        {
            //ExStart
            //ExFor:Document.LayoutOptions
            //ExFor:LayoutOptions
            //ExFor:LayoutOptions.ShowHiddenText
            //ExSummary:Shows how to hide text in a rendered output document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Assert.That(doc.LayoutOptions.ShowHiddenText, Is.False); //ExSkip

            // Insert hidden text, then specify whether we wish to omit it from a rendered document.
            builder.Writeln("This text is not hidden.");
            builder.Font.Hidden = true;
            builder.Writeln("This text is hidden.");

            doc.LayoutOptions.ShowHiddenText = showHiddenText;

            doc.Save(ArtifactsDir + "Document.LayoutOptionsHiddenText.pdf");
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void UsePdfDocumentForLayoutOptionsHiddenText(bool showHiddenText)
        {
            LayoutOptionsHiddenText(showHiddenText);

            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.LayoutOptionsHiddenText.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.That(textAbsorber.Text, Is.EqualTo(showHiddenText ?
                    $"This text is not hidden.{Environment.NewLine}This text is hidden." :
                    "This text is not hidden."));
        }

        [TestCase(false)]
        [TestCase(true)]
        public void LayoutOptionsParagraphMarks(bool showParagraphMarks)
        {
            //ExStart
            //ExFor:Document.LayoutOptions
            //ExFor:LayoutOptions
            //ExFor:LayoutOptions.ShowParagraphMarks
            //ExSummary:Shows how to show paragraph marks in a rendered output document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Assert.That(doc.LayoutOptions.ShowParagraphMarks, Is.False); //ExSkip

            // Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
            // with a pilcrow (¶) symbol when we render the document.
            builder.Writeln("Hello world!");
            builder.Writeln("Hello again!");

            doc.LayoutOptions.ShowParagraphMarks = showParagraphMarks;

            doc.Save(ArtifactsDir + "Document.LayoutOptionsParagraphMarks.pdf");
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void UsePdfDocumentForLayoutOptionsParagraphMarks(bool showParagraphMarks)
        {
            LayoutOptionsParagraphMarks(showParagraphMarks);

            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.LayoutOptionsParagraphMarks.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.That(textAbsorber.Text.Trim(), Is.EqualTo(showParagraphMarks ?
                    $"Hello world!¶{Environment.NewLine}Hello again!¶{Environment.NewLine}¶" :
                    $"Hello world!{Environment.NewLine}Hello again!"));
        }

        [Test]
        public void UpdatePageLayout()
        {
            //ExStart
            //ExFor:StyleCollection.Item(String)
            //ExFor:SectionCollection.Item(Int32)
            //ExFor:Document.UpdatePageLayout
            //ExFor:Margins
            //ExFor:PageSetup.Margins
            //ExSummary:Shows when to recalculate the page layout of the document.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Saving a document to PDF, to an image, or printing for the first time will automatically
            // cache the layout of the document within its pages.
            doc.Save(ArtifactsDir + "Document.UpdatePageLayout.1.pdf");

            // Modify the document in some way.
            doc.Styles["Normal"].Font.Size = 6;
            doc.Sections[0].PageSetup.Orientation = Aspose.Words.Orientation.Landscape;
            doc.Sections[0].PageSetup.Margins = Margins.Mirrored;

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

            Assert.That(doc.PackageCustomParts.Count, Is.EqualTo(2));

            // Clone the second part, then add the clone to the collection.
            CustomPart clonedPart = doc.PackageCustomParts[1].Clone();
            doc.PackageCustomParts.Add(clonedPart);
            TestDocPackageCustomParts(doc.PackageCustomParts); //ExSkip

            Assert.That(doc.PackageCustomParts.Count, Is.EqualTo(3));

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

            Assert.That(doc.PackageCustomParts.Count, Is.EqualTo(2));

            doc.PackageCustomParts.Clear();

            Assert.That(doc.PackageCustomParts.Count, Is.EqualTo(0));
            //ExEnd
        }

        private static void TestDocPackageCustomParts(CustomPartCollection parts)
        {
            Assert.That(parts.Count, Is.EqualTo(3));

            Assert.That(parts[0].Name, Is.EqualTo("/payload/payload_on_package.test"));
            Assert.That(parts[0].ContentType, Is.EqualTo("mytest/somedata"));
            Assert.That(parts[0].RelationshipType, Is.EqualTo("http://mytest.payload.internal"));
            Assert.That(parts[0].IsExternal, Is.EqualTo(false));
            Assert.That(parts[0].Data.Length, Is.EqualTo(18));

            Assert.That(parts[1].Name, Is.EqualTo("http://www.aspose.com/Images/aspose-logo.jpg"));
            Assert.That(parts[1].ContentType, Is.EqualTo(""));
            Assert.That(parts[1].RelationshipType, Is.EqualTo("http://mytest.payload.external"));
            Assert.That(parts[1].IsExternal, Is.EqualTo(true));
            Assert.That(parts[1].Data.Length, Is.EqualTo(0));

            Assert.That(parts[2].Name, Is.EqualTo("http://www.aspose.com/Images/aspose-logo.jpg"));
            Assert.That(parts[2].ContentType, Is.EqualTo(""));
            Assert.That(parts[2].RelationshipType, Is.EqualTo("http://mytest.payload.external"));
            Assert.That(parts[2].IsExternal, Is.EqualTo(true));
            Assert.That(parts[2].Data.Length, Is.EqualTo(0));
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
            Assert.That(doc.ShadeFormData, Is.True); //ExSkip

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
            Assert.That(doc.VersionsCount, Is.EqualTo(4));

            doc.Save(ArtifactsDir + "Document.VersionsCount.doc");
            doc = new Document(ArtifactsDir + "Document.VersionsCount.doc");

            Assert.That(doc.VersionsCount, Is.EqualTo(0));
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
            Assert.That(doc.WriteProtection.IsWriteProtected, Is.False); //ExSkip
            Assert.That(doc.WriteProtection.ReadOnlyRecommended, Is.False); //ExSkip

            // Enter a password up to 15 characters in length, and then verify the document's protection status.
            doc.WriteProtection.SetPassword("MyPassword");
            doc.WriteProtection.ReadOnlyRecommended = true;

            Assert.That(doc.WriteProtection.IsWriteProtected, Is.True);
            Assert.That(doc.WriteProtection.ValidatePassword("MyPassword"), Is.True);

            // Protection does not prevent the document from being edited programmatically, nor does it encrypt the contents.
            doc.Save(ArtifactsDir + "Document.WriteProtection.docx");
            doc = new Document(ArtifactsDir + "Document.WriteProtection.docx");

            Assert.That(doc.WriteProtection.IsWriteProtected, Is.True);

            builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln("Writing text in a protected document.");

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello world! This document is protected." +
                            "\rWriting text in a protected document."));
            //ExEnd
            Assert.That(doc.WriteProtection.ReadOnlyRecommended, Is.True);
            Assert.That(doc.WriteProtection.ValidatePassword("MyPassword"), Is.True);
            Assert.That(doc.WriteProtection.ValidatePassword("wrongpassword"), Is.False);
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

            Assert.That(doc.RemovePersonalInformation, Is.EqualTo(saveWithoutPersonalInfo));
            Assert.That(doc.BuiltInDocumentProperties.Author, Is.EqualTo("John Doe"));
            Assert.That(doc.BuiltInDocumentProperties.Company, Is.EqualTo("Placeholder Inc."));
            Assert.That(doc.Revisions[0].Author, Is.EqualTo("John Doe"));
            //ExEnd
        }

        [Test]
        public void ShowComments()
        {
            //ExStart
            //ExFor:LayoutOptions.CommentDisplayMode
            //ExFor:CommentDisplayMode
            //ExSummary:Shows how to show comments when saving a document to a rendered format.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello world!");

            Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            comment.SetText("My comment.");
            builder.CurrentParagraph.AppendChild(comment);

            // ShowInAnnotations is only available in Pdf1.7 and Pdf1.5 formats.
            // In other formats, it will work similarly to Hide.
            doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

            doc.Save(ArtifactsDir + "Document.ShowCommentsInAnnotations.pdf");

            // Note that it's required to rebuild the document page layout (via Document.UpdatePageLayout() method)
            // after changing the Document.LayoutOptions values.
            doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;
            doc.UpdatePageLayout();

            doc.Save(ArtifactsDir + "Document.ShowCommentsInBalloons.pdf");
            //ExEnd
        }

        [Test]
        public void UsePdfDocumentForShowComments()
        {
            ShowComments();

            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Document.ShowCommentsInBalloons.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.That(textAbsorber.Text, Is.EqualTo("Hello world!                                                                    Commented [J.D.1]:  My comment."));
        }

        [Test]
        public void CopyTemplateStylesViaDocument()
        {
            //ExStart
            //ExFor:Document.CopyStylesFromTemplate(Document)
            //ExSummary:Shows how to copies styles from the template to a document via Document.
            Document template = new Document(MyDir + "Rendering.docx");
            Document target = new Document(MyDir + "Document.docx");

            Assert.That(template.Styles.Count, Is.EqualTo(18)); //ExSkip
            Assert.That(target.Styles.Count, Is.EqualTo(12)); //ExSkip

            target.CopyStylesFromTemplate(template);
            Assert.That(target.Styles.Count, Is.EqualTo(22)); //ExSkip

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

            Assert.That(template.Styles.Count, Is.EqualTo(7));

            // Create a document which we will copy the styles to.
            Document target = new Document();

            // Create a style with the same name as a style from the template document and add it to the target document.
            style = target.Styles.Add(StyleType.Paragraph, "TemplateStyle3");
            style.Font.Name = "Calibri";
            style.Font.Color = Color.Orange;

            Assert.That(target.Styles.Count, Is.EqualTo(5));

            // There are two ways of calling the method to copy all the styles from one document to another.
            // 1 -  Passing the template document object:
            target.CopyStylesFromTemplate(template);

            // Copying styles adds all styles from the template document to the target
            // and overwrites existing styles with the same name.
            Assert.That(target.Styles.Count, Is.EqualTo(7));

            Assert.That(target.Styles["TemplateStyle3"].Font.Name, Is.EqualTo("Courier New"));
            Assert.That(target.Styles["TemplateStyle3"].Font.Color.ToArgb(), Is.EqualTo(Color.RoyalBlue.ToArgb()));

            // 2 -  Passing the local system filename of a template document:
            target.CopyStylesFromTemplate(MyDir + "Rendering.docx");

            Assert.That(target.Styles.Count, Is.EqualTo(21));
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
            Assert.That(vbaProject.IsSigned, Is.True); //ExSkip
            Console.WriteLine(vbaProject.IsSigned
                ? $"Project name: {vbaProject.Name} signed; Project code page: {vbaProject.CodePage}; Modules count: {vbaProject.Modules.Count()}\n"
                : $"Project name: {vbaProject.Name} not signed; Project code page: {vbaProject.CodePage}; Modules count: {vbaProject.Modules.Count()}\n");

            VbaModuleCollection vbaModules = doc.VbaProject.Modules;

            Assert.That(3, Is.EqualTo(vbaModules.Count()));

            foreach (VbaModule module in vbaModules)
                Console.WriteLine($"Module name: {module.Name};\nModule code:\n{module.SourceCode}\n");

            // Set new source code for VBA module. You can access VBA modules in the collection either by index or by name.
            vbaModules[0].SourceCode = "Your VBA code...";
            vbaModules["Module1"].SourceCode = "Your VBA code...";

            // Remove a module from the collection.
            vbaModules.Remove(vbaModules[2]);
            //ExEnd

            Assert.That(vbaProject.Name, Is.EqualTo("AsposeVBAtest"));
            Assert.That(vbaProject.Modules.Count(), Is.EqualTo(2));
            Assert.That(vbaProject.CodePage, Is.EqualTo(1251));
            Assert.That(vbaProject.IsSigned, Is.False);

            Assert.That(vbaModules[0].Name, Is.EqualTo("ThisDocument"));
            Assert.That(vbaModules[0].SourceCode, Is.EqualTo("Your VBA code..."));

            Assert.That(vbaModules[1].Name, Is.EqualTo("Module1"));
            Assert.That(vbaModules[1].SourceCode, Is.EqualTo("Your VBA code..."));
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

            Assert.That(parameters.ContentType, Is.EqualTo("application/msword"));

            // This property changes depending on the save format.
            parameters = doc.Save(ArtifactsDir + "Document.SaveOutputParameters.pdf");

            Assert.That(parameters.ContentType, Is.EqualTo("application/pdf"));
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
            Assert.That(subDocuments.Count, Is.EqualTo(1)); //ExSkip

            // This node serves as a reference to an external document, and its contents cannot be accessed.
            SubDocument subDocument = (SubDocument)subDocuments[0];

            Assert.That(subDocument.IsComposite, Is.False);
            //ExEnd
        }

        [Test]
        public void CreateWebExtension()
        {
            //ExStart
            //ExFor:BaseWebExtensionCollection`1.Add(`0)
            //ExFor:BaseWebExtensionCollection`1.Clear
            //ExFor:Document.WebExtensionTaskPanes
            //ExFor:TaskPane
            //ExFor:TaskPane.DockState
            //ExFor:TaskPane.IsVisible
            //ExFor:TaskPane.Width
            //ExFor:TaskPane.IsLocked
            //ExFor:TaskPane.WebExtension
            //ExFor:TaskPane.Row
            //ExFor:WebExtension
            //ExFor:WebExtension.Id
            //ExFor:WebExtension.AlternateReferences
            //ExFor:WebExtension.Reference
            //ExFor:WebExtension.Properties
            //ExFor:WebExtension.Bindings
            //ExFor:WebExtension.IsFrozen
            //ExFor:WebExtensionReference
            //ExFor:WebExtensionReference.Id
            //ExFor:WebExtensionReference.Version
            //ExFor:WebExtensionReference.StoreType
            //ExFor:WebExtensionReference.Store
            //ExFor:WebExtensionPropertyCollection
            //ExFor:WebExtensionBindingCollection
            //ExFor:WebExtensionProperty.#ctor(String, String)
            //ExFor:WebExtensionProperty.Name
            //ExFor:WebExtensionProperty.Value
            //ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
            //ExFor:WebExtensionStoreType
            //ExFor:WebExtensionBindingType
            //ExFor:TaskPaneDockState
            //ExFor:TaskPaneCollection
            //ExFor:WebExtensionBinding.Id
            //ExFor:WebExtensionBinding.AppRef
            //ExFor:WebExtensionBinding.BindingType
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

            Assert.That(doc.WebExtensionTaskPanes.Count, Is.EqualTo(0));

            doc = new Document(ArtifactsDir + "Document.WebExtension.docx");
            
            myScriptTaskPane = doc.WebExtensionTaskPanes[0];
            Assert.That(myScriptTaskPane.DockState, Is.EqualTo(TaskPaneDockState.Right));
            Assert.That(myScriptTaskPane.IsVisible, Is.True);
            Assert.That(myScriptTaskPane.Width, Is.EqualTo(300.0d));
            Assert.That(myScriptTaskPane.IsLocked, Is.True);
            Assert.That(myScriptTaskPane.Row, Is.EqualTo(1));

            webExtension = myScriptTaskPane.WebExtension;
            Assert.That(webExtension.Id, Is.EqualTo(string.Empty));

            Assert.That(webExtension.Reference.Id, Is.EqualTo("WA104380646"));
            Assert.That(webExtension.Reference.Version, Is.EqualTo("1.0.0.0"));
            Assert.That(webExtension.Reference.StoreType, Is.EqualTo(WebExtensionStoreType.OMEX));
            Assert.That(webExtension.Reference.Store, Is.EqualTo(CultureInfo.CurrentCulture.Name));
            Assert.That(webExtension.AlternateReferences.Count, Is.EqualTo(0));

            Assert.That(webExtension.Properties[0].Name, Is.EqualTo("MyScript"));
            Assert.That(webExtension.Properties[0].Value, Is.EqualTo("MyScript Math Sample"));

            Assert.That(webExtension.Bindings[0].Id, Is.EqualTo("MyScript"));
            Assert.That(webExtension.Bindings[0].BindingType, Is.EqualTo(WebExtensionBindingType.Text));
            Assert.That(webExtension.Bindings[0].AppRef, Is.EqualTo("104380646"));

            Assert.That(webExtension.IsFrozen, Is.False);
            //ExEnd
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

            Assert.That(doc.WebExtensionTaskPanes.Count, Is.EqualTo(1));

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

            Assert.That(doc.WebExtensionTaskPanes.Count, Is.EqualTo(0));
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
            //ExFor:Document.Watermark
            //ExFor:Watermark
            //ExFor:Watermark.SetText(String)
            //ExFor:Watermark.SetText(String, TextWatermarkOptions)
            //ExFor:Watermark.Remove
            //ExFor:TextWatermarkOptions
            //ExFor:TextWatermarkOptions.FontFamily
            //ExFor:TextWatermarkOptions.FontSize
            //ExFor:TextWatermarkOptions.Color
            //ExFor:TextWatermarkOptions.Layout
            //ExFor:TextWatermarkOptions.IsSemitrasparent
            //ExFor:WatermarkLayout
            //ExFor:WatermarkType
            //ExFor:Watermark.Type
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

            Assert.That(doc.Watermark.Type, Is.EqualTo(WatermarkType.Text));
        }

        [Test]
        public void ImageWatermark()
        {
            //ExStart
            //ExFor:Watermark.SetImage(Image)
            //ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
            //ExFor:Watermark.SetImage(String, ImageWatermarkOptions)
            //ExFor:ImageWatermarkOptions
            //ExFor:ImageWatermarkOptions.Scale
            //ExFor:ImageWatermarkOptions.IsWashout
            //ExSummary:Shows how to create a watermark from an image in the local file system.
            Document doc = new Document();

            // Modify the image watermark's appearance with an ImageWatermarkOptions object,
            // then pass it while creating a watermark from an image file.
            ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
            imageWatermarkOptions.Scale = 5;
            imageWatermarkOptions.IsWashout = false;

#if NET461_OR_GREATER || JAVA || CPLUSPLUS
            // We have a different options to insert image.
            // Use on of the following methods to add image watermark.
            doc.Watermark.SetImage(Image.FromFile(ImageDir + "Logo.jpg"));

            doc.Watermark.SetImage(Image.FromFile(ImageDir + "Logo.jpg"), imageWatermarkOptions);

            doc.Watermark.SetImage(ImageDir + "Logo.jpg", imageWatermarkOptions);

#elif NET6_0_OR_GREATER
            using (SKBitmap image = SKBitmap.Decode(ImageDir + "Logo.jpg"))
                doc.Watermark.SetImage(image, imageWatermarkOptions);
#endif

            doc.Save(ArtifactsDir + "Document.ImageWatermark.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Document.ImageWatermark.docx");
            Assert.That(doc.Watermark.Type, Is.EqualTo(WatermarkType.Image));
        }

        [Test]
        public void ImageWatermarkStream()
        {
            //ExStart:ImageWatermarkStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermark.SetImage(Stream, ImageWatermarkOptions)
            //ExSummary:Shows how to create a watermark from an image stream.
            Document doc = new Document();

            // Modify the image watermark's appearance with an ImageWatermarkOptions object,
            // then pass it while creating a watermark from an image file.
            ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
            imageWatermarkOptions.Scale = 5;

            using (FileStream imageStream = new FileStream(ImageDir + "Logo.jpg", FileMode.Open, FileAccess.Read))
                doc.Watermark.SetImage(imageStream, imageWatermarkOptions);

            doc.Save(ArtifactsDir + "Document.ImageWatermarkStream.docx");
            //ExEnd:ImageWatermarkStream

            doc = new Document(ArtifactsDir + "Document.ImageWatermarkStream.docx");
            Assert.That(doc.Watermark.Type, Is.EqualTo(WatermarkType.Image));
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

            Assert.That(doc.ShowGrammaticalErrors, Is.EqualTo(showErrors));
            Assert.That(doc.ShowSpellingErrors, Is.EqualTo(showErrors));
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
            Assert.That(2, Is.EqualTo(doc.PageCount));
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

        [Test]
        public void Frameset()
        {
            //ExStart
            //ExFor:Document.Frameset
            //ExFor:Frameset
            //ExFor:Frameset.FrameDefaultUrl
            //ExFor:Frameset.IsFrameLinkToFile
            //ExFor:Frameset.ChildFramesets
            //ExFor:FramesetCollection
            //ExFor:FramesetCollection.Count
            //ExFor:FramesetCollection.Item(Int32)
            //ExSummary:Shows how to access frames on-page.
            // Document contains several frames with links to other documents.
            Document doc = new Document(MyDir + "Frameset.docx");

            Assert.That(doc.Frameset.ChildFramesets.Count, Is.EqualTo(3));
            // We can check the default URL (a web page URL or local document) or if the frame is an external resource.
            Assert.That(doc.Frameset.ChildFramesets[0].ChildFramesets[0].FrameDefaultUrl, Is.EqualTo("https://file-examples-com.github.io/uploads/2017/02/file-sample_100kB.docx"));
            Assert.That(doc.Frameset.ChildFramesets[0].ChildFramesets[0].IsFrameLinkToFile, Is.True);

            Assert.That(doc.Frameset.ChildFramesets[1].FrameDefaultUrl, Is.EqualTo("Document.docx"));
            Assert.That(doc.Frameset.ChildFramesets[1].IsFrameLinkToFile, Is.False);

            // Change properties for one of our frames.
            doc.Frameset.ChildFramesets[0].ChildFramesets[0].FrameDefaultUrl =
                "https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx";
            doc.Frameset.ChildFramesets[0].ChildFramesets[0].IsFrameLinkToFile = false;
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.Frameset.ChildFramesets[0].ChildFramesets[0].FrameDefaultUrl, Is.EqualTo("https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx"));
            Assert.That(doc.Frameset.ChildFramesets[0].ChildFramesets[0].IsFrameLinkToFile, Is.False);
        }

        [Test]
        public void OpenAzw()
        {
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Azw3 document.azw3");
            Assert.That(LoadFormat.Azw3, Is.EqualTo(info.LoadFormat));

            Document doc = new Document(MyDir + "Azw3 document.azw3");
            Assert.That(doc.GetText().Contains("Hachette Book Group USA"), Is.True);
        }

        [Test]
        public void OpenEpub()
        {
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Epub document.epub");
            Assert.That(LoadFormat.Epub, Is.EqualTo(info.LoadFormat));

            Document doc = new Document(MyDir + "Epub document.epub");
            Assert.That(doc.GetText().Contains("Down the Rabbit-Hole"), Is.True);
        }

        [Test]
        public void OpenXml()
        {
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Mail merge data - Customers.xml");
            Assert.That(LoadFormat.Xml, Is.EqualTo(info.LoadFormat));

            Document doc = new Document(MyDir + "Mail merge data - Purchase order.xml");
            Assert.That(doc.GetText().Contains("Ellen Adams\r123 Maple Street"), Is.True);
        }

        [Test]
        public void MoveToStructuredDocumentTag()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToStructuredDocumentTag(int, int)
            //ExFor:DocumentBuilder.MoveToStructuredDocumentTag(StructuredDocumentTag, int)
            //ExFor:DocumentBuilder.IsAtEndOfStructuredDocumentTag
            //ExFor:DocumentBuilder.CurrentStructuredDocumentTag
            //ExSummary:Shows how to move cursor of DocumentBuilder inside a structured document tag.
            Document doc = new Document(MyDir + "Structured document tags.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // There is a several ways to move the cursor:
            // 1 -  Move to the first character of structured document tag by index.
            builder.MoveToStructuredDocumentTag(1, 1);

            // 2 -  Move to the first character of structured document tag by object.
            StructuredDocumentTag tag = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 2, true);
            builder.MoveToStructuredDocumentTag(tag, 1);
            builder.Write(" New text.");

            Assert.That(tag.GetText().Trim(), Is.EqualTo("R New text.ichText"));

            // 3 -  Move to the end of the second structured document tag.
            builder.MoveToStructuredDocumentTag(1, -1);
            Assert.That(builder.IsAtEndOfStructuredDocumentTag, Is.True);

            // Get currently selected structured document tag.
            builder.CurrentStructuredDocumentTag.Color = Color.Green;

            doc.Save(ArtifactsDir + "Document.MoveToStructuredDocumentTag.docx");
            //ExEnd
        }

        [Test]
        public void IncludeTextboxesFootnotesEndnotesInStat()
        {
            //ExStart
            //ExFor:Document.IncludeTextboxesFootnotesEndnotesInStat
            //ExSummary: Shows how to include or exclude textboxes, footnotes and endnotes from word count statistics.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Lorem ipsum");
            builder.InsertFootnote(FootnoteType.Footnote, "sit amet");

            // By default option is set to 'false'.
            doc.UpdateWordCount();
            // Words count without textboxes, footnotes and endnotes.
            Assert.That(doc.BuiltInDocumentProperties.Words, Is.EqualTo(2));

            doc.IncludeTextboxesFootnotesEndnotesInStat = true;
            doc.UpdateWordCount();
            // Words count with textboxes, footnotes and endnotes.
            Assert.That(doc.BuiltInDocumentProperties.Words, Is.EqualTo(4));
            //ExEnd
        }

        [Test]
        public void SetJustificationMode()
        {
            //ExStart
            //ExFor:Document.JustificationMode
            //ExFor:JustificationMode
            //ExSummary:Shows how to manage character spacing control.
            Document doc = new Document(MyDir + "Document.docx");

            JustificationMode justificationMode = doc.JustificationMode;
            if (justificationMode == JustificationMode.Expand)
                doc.JustificationMode = JustificationMode.Compress;

            doc.Save(ArtifactsDir + "Document.SetJustificationMode.docx");
            //ExEnd
        }

        [Test]
        public void PageIsInColor()
        {
            //ExStart
            //ExFor:PageInfo.Colored
            //ExFor:Document.GetPageInfo(Int32)
            //ExSummary:Shows how to check whether the page is in color or not.
            Document doc = new Document(MyDir + "Document.docx");

            // Check that the first page of the document is not colored.
            Assert.That(doc.GetPageInfo(0).Colored, Is.False);
            //ExEnd
        }

        [Test]
        public void InsertDocumentInline()
        {
            //ExStart:InsertDocumentInline
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:DocumentBuilder.InsertDocumentInline(Document, ImportFormatMode, ImportFormatOptions)
            //ExSummary:Shows how to insert a document inline at the cursor position.
            DocumentBuilder srcDoc = new DocumentBuilder();
            srcDoc.Write("[src content]");

            // Create destination document.
            DocumentBuilder dstDoc = new DocumentBuilder();
            dstDoc.Write("Before ");
            dstDoc.InsertNode(new BookmarkStart(dstDoc.Document, "src_place"));
            dstDoc.InsertNode(new BookmarkEnd(dstDoc.Document, "src_place"));
            dstDoc.Write(" after");

            Assert.That(dstDoc.Document.GetText().TrimEnd(), Is.EqualTo("Before  after"));

            // Insert source document into destination inline.
            dstDoc.MoveToBookmark("src_place");
            dstDoc.InsertDocumentInline(srcDoc.Document, ImportFormatMode.UseDestinationStyles, new ImportFormatOptions());

            Assert.That(dstDoc.Document.GetText().TrimEnd(), Is.EqualTo("Before [src content] after"));
            //ExEnd:InsertDocumentInline
        }

        [TestCase(SaveFormat.Doc)]
        [TestCase(SaveFormat.Dot)]
        [TestCase(SaveFormat.Docx)]
        [TestCase(SaveFormat.Docm)]
        [TestCase(SaveFormat.Dotx)]
        [TestCase(SaveFormat.Dotm)]
        [TestCase(SaveFormat.FlatOpc)]
        [TestCase(SaveFormat.FlatOpcMacroEnabled)]
        [TestCase(SaveFormat.FlatOpcTemplate)]
        [TestCase(SaveFormat.FlatOpcTemplateMacroEnabled)]
        [TestCase(SaveFormat.Rtf)]
        [TestCase(SaveFormat.WordML)]
        [TestCase(SaveFormat.Pdf)]
        [TestCase(SaveFormat.Xps)]
        [TestCase(SaveFormat.XamlFixed)]
        [TestCase(SaveFormat.Svg)]
        [TestCase(SaveFormat.HtmlFixed)]
        [TestCase(SaveFormat.OpenXps)]
        [TestCase(SaveFormat.Ps)]
        [TestCase(SaveFormat.Pcl)]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        [TestCase(SaveFormat.Azw3)]
        [TestCase(SaveFormat.Mobi)]
        [TestCase(SaveFormat.Odt)]
        [TestCase(SaveFormat.Ott)]
        [TestCase(SaveFormat.Text)]
        [TestCase(SaveFormat.XamlFlow)]
        [TestCase(SaveFormat.XamlFlowPack)]
        [TestCase(SaveFormat.Markdown)]
        [TestCase(SaveFormat.Xlsx)]
        [TestCase(SaveFormat.Tiff)]
        [TestCase(SaveFormat.Png)]
        [TestCase(SaveFormat.Bmp)]
        [TestCase(SaveFormat.Emf)]
        [TestCase(SaveFormat.Jpeg)]
        [TestCase(SaveFormat.Gif)]
        [TestCase(SaveFormat.Eps)]
        public void SaveDocumentToStream(SaveFormat saveFormat)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Lorem ipsum");

            using (Stream stream = new MemoryStream())
            {
                if (saveFormat == SaveFormat.HtmlFixed)
                {
                    HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
                    saveOptions.ExportEmbeddedCss = true;
                    saveOptions.ExportEmbeddedFonts = true;
                    saveOptions.SaveFormat = saveFormat;

                    doc.Save(stream, saveOptions);
                }
                else if (saveFormat == SaveFormat.XamlFixed)
                {
                    XamlFixedSaveOptions saveOptions = new XamlFixedSaveOptions();
                    saveOptions.ResourcesFolder = ArtifactsDir;
                    saveOptions.SaveFormat = saveFormat;

                    doc.Save(stream, saveOptions);
                }
                else
                    doc.Save(stream, saveFormat);
            }
        }

        [Test]
        public void HasMacros()
        {
            //ExStart:HasMacros
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:FileFormatInfo.HasMacros
            //ExSummary:Shows how to check VBA macro presence without loading document.
            FileFormatInfo fileFormatInfo = FileFormatUtil.DetectFileFormat(MyDir + "Macro.docm");
            Assert.That(fileFormatInfo.HasMacros, Is.True);
            //ExEnd:HasMacros
        }

        [Test]
        public void PunctuationKerning()
        {
            //ExStart
            //ExFor:Document.PunctuationKerning
            //ExSummary:Shows how to work with kerning applies to both Latin text and punctuation.
            Document doc = new Document(MyDir + "Document.docx");
            Assert.That(doc.PunctuationKerning, Is.True);
            //ExEnd
        }

        [Test]
        public void RemoveBlankPages()
        {
            //ExStart
            //ExFor:Document.RemoveBlankPages
            //ExSummary:Shows how to remove blank pages from the document.
            Document doc = new Document(MyDir + "Blank pages.docx");
            Assert.That(doc.PageCount, Is.EqualTo(2));
            doc.RemoveBlankPages();
            doc.UpdatePageLayout();
            Assert.That(doc.PageCount, Is.EqualTo(1));
            //ExEnd
        }
    }
}
