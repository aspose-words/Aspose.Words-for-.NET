﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Pdf.Text;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Loading;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlLoadOptions : ApiExampleBase
    {
        [TestCase(true, Category = "SkipMono")]
        [TestCase(false, Category = "SkipMono")]
        public void SupportVml(bool supportVml)
        {
            //ExStart
            //ExFor:HtmlLoadOptions
            //ExFor:HtmlLoadOptions.#ctor
            //ExFor:HtmlLoadOptions.SupportVml
            //ExSummary:Shows how to support conditional comments while loading an HTML document.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // If the value is true, then we take VML code into account while parsing the loaded document.
            loadOptions.SupportVml = supportVml;

            // This document contains a JPEG image within "<!--[if gte vml 1]>" tags,
            // and a different PNG image within "<![if !vml]>" tags.
            // If we set the "SupportVml" flag to "true", then Aspose.Words will load the JPEG.
            // If we set this flag to "false", then Aspose.Words will only load the PNG.
            Document doc = new Document(MyDir + "VML conditional.htm", loadOptions);

            if (supportVml)
                Assert.That(((Shape)doc.GetChild(NodeType.Shape, 0, true)).ImageData.ImageType, Is.EqualTo(ImageType.Jpeg));
            else
                Assert.That(((Shape)doc.GetChild(NodeType.Shape, 0, true)).ImageData.ImageType, Is.EqualTo(ImageType.Png));
            //ExEnd

            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            if (supportVml)
                TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, imageShape);
            else
                TestUtil.VerifyImageInShape(400, 400, ImageType.Png, imageShape);
        }

        //ExStart
        //ExFor:HtmlLoadOptions.WebRequestTimeout
        //ExSummary:Shows how to set a time limit for web requests when loading a document with external resources linked by URLs.
        [Test] //ExSkip
        public void WebRequestTimeout()
        {
            // Create a new HtmlLoadOptions object and verify its timeout threshold for a web request.
            HtmlLoadOptions options = new HtmlLoadOptions();

            // When loading an Html document with resources externally linked by a web address URL,
            // Aspose.Words will abort web requests that fail to fetch the resources within this time limit, in milliseconds.
            Assert.That(options.WebRequestTimeout, Is.EqualTo(100000));

            // Set a WarningCallback that will record all warnings that occur during loading.
            ListDocumentWarnings warningCallback = new ListDocumentWarnings();
            options.WarningCallback = warningCallback;

            // Load such a document and verify that a shape with image data has been created.
            // This linked image will require a web request to load, which will have to complete within our time limit.
            string html = $@"
                <html>
                    <img src=""{ImageUrl}"" alt=""Aspose logo"" style=""width:400px;height:400px;"">
                </html>
            ";

            // Set an unreasonable timeout limit and try load the document again.
            options.WebRequestTimeout = 0;
            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), options);
            Assert.That(warningCallback.Warnings().Count, Is.EqualTo(2));

            // A web request that fails to obtain an image within the time limit will still produce an image.
            // However, the image will be the red 'x' that commonly signifies missing images.
            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.That(imageShape.ImageData.ImageBytes.Length, Is.EqualTo(924));

            // We can also configure a custom callback to pick up any warnings from timed out web requests.
            Assert.That(warningCallback.Warnings()[0].Source, Is.EqualTo(WarningSource.Html));
            Assert.That(warningCallback.Warnings()[0].WarningType, Is.EqualTo(WarningType.DataLoss));
            Assert.That(warningCallback.Warnings()[0].Description, Is.EqualTo($"Couldn't load a resource from \'{ImageUrl}\'."));

            Assert.That(warningCallback.Warnings()[1].Source, Is.EqualTo(WarningSource.Html));
            Assert.That(warningCallback.Warnings()[1].WarningType, Is.EqualTo(WarningType.DataLoss));
            Assert.That(warningCallback.Warnings()[1].Description, Is.EqualTo("Image has been replaced with a placeholder."));

            doc.Save(ArtifactsDir + "HtmlLoadOptions.WebRequestTimeout.docx");
        }

        /// <summary>
        /// Stores all warnings that occur during a document loading operation in a List.
        /// </summary>
        private class ListDocumentWarnings : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                mWarnings.Add(info);
            }

            public List<WarningInfo> Warnings() { 
                return mWarnings;
            }

            private readonly List<WarningInfo> mWarnings = new List<WarningInfo>();
        }
        //ExEnd

        [Test]
        public void LoadHtmlFixed()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions { SaveFormat = SaveFormat.HtmlFixed };

            doc.Save(ArtifactsDir + "HtmlLoadOptions.Fixed.html", saveOptions);

            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            ListDocumentWarnings warningCallback = new ListDocumentWarnings();
            loadOptions.WarningCallback = warningCallback;

            doc = new Document(ArtifactsDir + "HtmlLoadOptions.Fixed.html", loadOptions);
            Assert.That(warningCallback.Warnings().Count, Is.EqualTo(1));

            Assert.That(warningCallback.Warnings()[0].Source, Is.EqualTo(WarningSource.Html));
            Assert.That(warningCallback.Warnings()[0].WarningType, Is.EqualTo(WarningType.MajorFormattingLoss));
            Assert.That(warningCallback.Warnings()[0].Description, Is.EqualTo("The document is fixed-page HTML. Its structure may not be loaded correctly."));
        }

        [Test]
        public void EncryptedHtml()
        {
            //ExStart
            //ExFor:HtmlLoadOptions.#ctor(String)
            //ExSummary:Shows how to encrypt an Html document, and then open it using a password.
            // Create and sign an encrypted HTML document from an encrypted .docx.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            SignOptions signOptions = new SignOptions
            {
                Comments = "Comment",
                SignTime = DateTime.Now,
                DecryptionPassword = "docPassword"
            };

            string inputFileName = MyDir + "Encrypted.docx";
            string outputFileName = ArtifactsDir + "HtmlLoadOptions.EncryptedHtml.html";
            DigitalSignatureUtil.Sign(inputFileName, outputFileName, certificateHolder, signOptions);

            // To load and read this document, we will need to pass its decryption
            // password using a HtmlLoadOptions object.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions("docPassword");

            Assert.That(loadOptions.Password, Is.EqualTo(signOptions.DecryptionPassword));

            Document doc = new Document(outputFileName, loadOptions);

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Test encrypted document."));
            //ExEnd
        }

        [Test]
        public void BaseUri()
        {
            //ExStart
            //ExFor:HtmlLoadOptions.#ctor(LoadFormat,String,String)
            //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
            //ExFor:LoadOptions.LoadFormat
            //ExFor:LoadFormat
            //ExSummary:Shows how to specify a base URI when opening an html document.
            // Suppose we want to load an .html document that contains an image linked by a relative URI
            // while the image is in a different location. In that case, we will need to resolve the relative URI into an absolute one.
            // We can provide a base URI using an HtmlLoadOptions object. 
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html, "", ImageDir);

            Assert.That(loadOptions.LoadFormat, Is.EqualTo(LoadFormat.Html));

            Document doc = new Document(MyDir + "Missing image.html", loadOptions);

            // While the image was broken in the input .html, our custom base URI helped us repair the link.
            Shape imageShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
            Assert.That(imageShape.IsImage, Is.True);

            // This output document will display the image that was missing.
            doc.Save(ArtifactsDir + "HtmlLoadOptions.BaseUri.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "HtmlLoadOptions.BaseUri.docx");

            Assert.That(((Shape)doc.GetChild(NodeType.Shape, 0, true)).ImageData.ImageBytes.Length > 0, Is.True);
        }

        [Test]
        public void GetSelectAsSdt()
        {
            //ExStart
            //ExFor:HtmlLoadOptions.PreferredControlType
            //ExFor:HtmlControlType
            //ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
            const string html = @"
                <html>
                    <select name='ComboBox' size='1'>
                        <option value='val1'>item1</option>
                        <option value='val2'></option>
                    </select>
                </html>
            ";

            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
            htmlLoadOptions.PreferredControlType = HtmlControlType.StructuredDocumentTag;

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), htmlLoadOptions);
            NodeCollection nodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            StructuredDocumentTag tag = (StructuredDocumentTag) nodes[0];
            //ExEnd

            Assert.That(tag.ListItems.Count, Is.EqualTo(2));

            Assert.That(tag.ListItems[0].Value, Is.EqualTo("val1"));
            Assert.That(tag.ListItems[1].Value, Is.EqualTo("val2"));
        }

        [Test]
        public void GetInputAsFormField()
        {
            const string html = @"
                <html>
                    <input type='text' value='Input value text' />
                </html>
            ";

            // By default, "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField".
            // So, we do not set this value.
            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), htmlLoadOptions);
            NodeCollection nodes = doc.GetChildNodes(NodeType.FormField, true);

            Assert.That(nodes.Count, Is.EqualTo(1));

            FormField formField = (FormField) nodes[0];
            Assert.That(formField.Result, Is.EqualTo("Input value text"));
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IgnoreNoscriptElements(bool ignoreNoscriptElements)
        {
            //ExStart
            //ExFor:HtmlLoadOptions.IgnoreNoscriptElements
            //ExSummary:Shows how to ignore <noscript> HTML elements.
            const string html = @"
                <html>
                  <head>
                    <title>NOSCRIPT</title>
                      <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
                      <script type=""text/javascript"">
                        alert(""Hello, world!"");
                      </script>
                  </head>
                <body>
                  <noscript><p>Your browser does not support JavaScript!</p></noscript>
                </body>
                </html>";

            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();
            htmlLoadOptions.IgnoreNoscriptElements = ignoreNoscriptElements;

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), htmlLoadOptions);
            doc.Save(ArtifactsDir + "HtmlLoadOptions.IgnoreNoscriptElements.pdf");
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void UsePdfDocumentForIgnoreNoscriptElements(bool ignoreNoscriptElements)
        {
            IgnoreNoscriptElements(ignoreNoscriptElements);

            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "HtmlLoadOptions.IgnoreNoscriptElements.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.That(textAbsorber.Text, Is.EqualTo(ignoreNoscriptElements ? "" : "Your browser does not support JavaScript!"));
        }

        [TestCase(BlockImportMode.Preserve)]
        [TestCase(BlockImportMode.Merge)]
        public void BlockImport(BlockImportMode blockImportMode)
        {
            //ExStart
            //ExFor:HtmlLoadOptions.BlockImportMode
            //ExFor:BlockImportMode
            //ExSummary:Shows how properties of block-level elements are imported from HTML-based documents.
            const string html = @"
            <html>
                <div style='border:dotted'>
                    <div style='border:solid'>
                        <p>paragraph 1</p>
                        <p>paragraph 2</p>
                    </div>
                </div>
            </html>";
            MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(html));

            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            // Set the new mode of import HTML block-level elements.
            loadOptions.BlockImportMode = blockImportMode;

            Document doc = new Document(stream, loadOptions);
            doc.Save(ArtifactsDir + "HtmlLoadOptions.BlockImport.docx");
            //ExEnd
        }

        [Test]
        public void FontFaceRules()
        {
            //ExStart:FontFaceRules
            //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
            //ExFor:HtmlLoadOptions.SupportFontFaceRules
            //ExSummary:Shows how to load declared "@font-face" rules.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportFontFaceRules = true;
            Document doc = new Document(MyDir + "Html with FontFace.html", loadOptions);

            Assert.That(doc.FontInfos[0].Name, Is.EqualTo("Squarish Sans CT Regular"));
            //ExEnd:FontFaceRules
        }
    }
}