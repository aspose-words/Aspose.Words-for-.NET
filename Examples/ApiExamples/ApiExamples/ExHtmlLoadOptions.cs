// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Markup;
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
                Assert.AreEqual(ImageType.Jpeg, ((Shape)doc.GetChild(NodeType.Shape, 0, true)).ImageData.ImageType);
            else
                Assert.AreEqual(ImageType.Png, ((Shape)doc.GetChild(NodeType.Shape, 0, true)).ImageData.ImageType);
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
            Assert.AreEqual(100000, options.WebRequestTimeout);

            // Set a WarningCallback that will record all warnings that occur during loading.
            ListDocumentWarnings warningCallback = new ListDocumentWarnings();
            options.WarningCallback = warningCallback;

            // Load such a document and verify that a shape with image data has been created.
            // This linked image will require a web request to load, which will have to complete within our time limit.
            string html = $@"
                <html>
                    <img src=""{AsposeLogoUrl}"" alt=""Aspose logo"" style=""width:400px;height:400px;"">
                </html>
            ";

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), options);
            Shape imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(7498, imageShape.ImageData.ImageBytes.Length);
            Assert.AreEqual(0, warningCallback.Warnings().Count);

            // Set an unreasonable timeout limit and try load the document again.
            options.WebRequestTimeout = 0;
            doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), options);

            // A web request that fails to obtain an image within the time limit will still produce an image.
            // However, the image will be the red 'x' that commonly signifies missing images.
            imageShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.AreEqual(924, imageShape.ImageData.ImageBytes.Length);

            // We can also configure a custom callback to pick up any warnings from timed out web requests.
            Assert.AreEqual(WarningSource.Html, warningCallback.Warnings()[0].Source);
            Assert.AreEqual(WarningType.DataLoss, warningCallback.Warnings()[0].WarningType);
            Assert.AreEqual($"Couldn't load a resource from \'{AsposeLogoUrl}\'.", warningCallback.Warnings()[0].Description);

            Assert.AreEqual(WarningSource.Html, warningCallback.Warnings()[1].Source);
            Assert.AreEqual(WarningType.DataLoss, warningCallback.Warnings()[1].WarningType);
            Assert.AreEqual("Image has been replaced with a placeholder.", warningCallback.Warnings()[1].Description);

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

            Assert.AreEqual(signOptions.DecryptionPassword, loadOptions.Password);

            Document doc = new Document(outputFileName, loadOptions);

            Assert.AreEqual("Test encrypted document.", doc.GetText().Trim());       
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

            Assert.AreEqual(LoadFormat.Html, loadOptions.LoadFormat);

            Document doc = new Document(MyDir + "Missing image.html", loadOptions);

            // While the image was broken in the input .html, our custom base URI helped us repair the link.
            Shape imageShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
            Assert.True(imageShape.IsImage);

            // This output document will display the image that was missing.
            doc.Save(ArtifactsDir + "HtmlLoadOptions.BaseUri.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "HtmlLoadOptions.BaseUri.docx");

            Assert.True(((Shape)doc.GetChild(NodeType.Shape, 0, true)).ImageData.ImageBytes.Length > 0);
        }

        [Test]
        public void GetSelectAsSdt()
        {
            //ExStart
            //ExFor:HtmlLoadOptions.PreferredControlType
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

            Assert.AreEqual(2, tag.ListItems.Count);

            Assert.AreEqual("val1", tag.ListItems[0].Value);
            Assert.AreEqual("val2", tag.ListItems[1].Value);
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

            Assert.AreEqual(1, nodes.Count);

            FormField formField = (FormField) nodes[0];
            Assert.AreEqual("Input value text", formField.Result);
        }
    }
}