// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
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
        [Test]
        public void SupportVml()
        {
            //ExStart
            //ExFor:HtmlLoadOptions.#ctor
            //ExFor:HtmlLoadOptions.SupportVml
            //ExFor:HtmlLoadOptions.WebRequestTimeout
            //ExSummary:Shows how to parse HTML document with conditional comments like "<!--[if gte vml 1]>" and "<![if !vml]>"
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // If value is true, then we parse "<!--[if gte vml 1]>", else parse "<![if !vml]>"
            loadOptions.SupportVml = true;

            // Wait for a response, when loading external resources
            loadOptions.WebRequestTimeout = 1000;

            Document doc = new Document(MyDir + "Shape.VmlAndDml.htm", loadOptions);
            doc.Save(ArtifactsDir + "Shape.VmlAndDml.docx");
            //ExEnd
        }

        [Test]
        public void EncryptedHtml()
        {
            //ExStart
            //ExFor:HtmlLoadOptions.#ctor(String)
            //ExSummary:Shows how to encrypt an Html document and then open it using a password.
            // Create and sign an encrypted html document from an encrypted .docx
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            SignOptions signOptions = new SignOptions
            {
                Comments = "Comment",
                SignTime = DateTime.Now,
                DecryptionPassword = "docPassword"
            };

            string inputFileName = MyDir + "Document.Encrypted.docx";
            string outputFileName = ArtifactsDir + "HtmlLoadOptions.EncryptedHtml.html";
            DigitalSignatureUtil.Sign(inputFileName, outputFileName, certificateHolder, signOptions);

            // This .html document will need a password to be decrypted, opened and have its contents accessed
            // The password is specified by HtmlLoadOptions.Password
            HtmlLoadOptions loadOptions = new HtmlLoadOptions("docPassword");
            Assert.AreEqual(signOptions.DecryptionPassword, loadOptions.Password);

            Document doc = new Document(outputFileName, loadOptions);
            Assert.AreEqual("Test signed document.", doc.GetText().Trim());       
            //ExEnd
        }

        [Test]
        public void BaseUri()
        {
            //ExStart
            //ExFor:HtmlLoadOptions.#ctor(LoadFormat,String,String)
            //ExSummary:Shows how to specify a base URI when opening an html document.
            // Create and sign an encrypted html document from an encrypted .docx
            // If we want to load an .html document which contains an image linked by a relative URI
            // while the image is in a different location, we will need to resolve the relative URI into an absolute one
            // by creating an HtmlLoadOptions and providing a base URI 
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html, "", MyDir + "/images/");

            Document doc = new Document(MyDir + "Document.OpenFromStreamWithBaseUri.html", loadOptions);

            // The image will be displayed correctly by the output document and
            doc.Save(ArtifactsDir + "Shape.BaseUri.docx");
        
            Shape imgShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
            Assert.True(imgShape.IsImage);

            imgShape.ImageData.Save(ArtifactsDir + "BaseUri.png");
            //ExEnd
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

            // By default "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField"
            // So, we do not set this value
            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions();

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), htmlLoadOptions);
            NodeCollection nodes = doc.GetChildNodes(NodeType.FormField, true);

            Assert.AreEqual(1, nodes.Count);

            FormField formField = (FormField) nodes[0];
            Assert.AreEqual("Input value text", formField.Result);
        }
    }
}