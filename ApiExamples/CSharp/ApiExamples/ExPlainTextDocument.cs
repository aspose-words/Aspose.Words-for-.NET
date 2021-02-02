// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExPlainTextDocument : ApiExampleBase
    {
        [Test]
        public void Load()
        {
            //ExStart
            //ExFor:PlainTextDocument
            //ExFor:PlainTextDocument.#ctor(String)
            //ExFor:PlainTextDocument.Text
            //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext.
            Document doc = new Document(); 
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            doc.Save(ArtifactsDir + "PlainTextDocument.Load.docx");

            PlainTextDocument plaintext = new PlainTextDocument(ArtifactsDir + "PlainTextDocument.Load.docx");

            Assert.AreEqual("Hello world!", plaintext.Text.Trim());
            //ExEnd
        }

        [Test]
        public void LoadFromStream()
        {
            //ExStart
            //ExFor:PlainTextDocument.#ctor(Stream)
            //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext using stream.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            doc.Save(ArtifactsDir + "PlainTextDocument.LoadFromStream.docx");

            using (FileStream stream = new FileStream(ArtifactsDir + "PlainTextDocument.LoadFromStream.docx", FileMode.Open))
            {
                PlainTextDocument plaintext = new PlainTextDocument(stream);

                Assert.AreEqual("Hello world!", plaintext.Text.Trim());
            }
            //ExEnd
        }

        [Test]
        public void LoadEncrypted()
        {
            //ExStart
            //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
            //ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "MyPassword";

            doc.Save(ArtifactsDir + "PlainTextDocument.LoadEncrypted.docx", saveOptions);

            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MyPassword";

            PlainTextDocument plaintext = new PlainTextDocument(ArtifactsDir + "PlainTextDocument.LoadEncrypted.docx", loadOptions);

            Assert.AreEqual("Hello world!", plaintext.Text.Trim());
            //ExEnd
        }

        [Test]
        public void LoadEncryptedUsingStream()
        {
            //ExStart
            //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
            //ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext using stream.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "MyPassword";

            doc.Save(ArtifactsDir + "PlainTextDocument.LoadFromStreamWithOptions.docx", saveOptions);

            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "MyPassword";

            using (FileStream stream = new FileStream(ArtifactsDir + "PlainTextDocument.LoadFromStreamWithOptions.docx", FileMode.Open))
            {
                PlainTextDocument plaintext = new PlainTextDocument(stream, loadOptions);

                Assert.AreEqual("Hello world!", plaintext.Text.Trim());
            }
            //ExEnd
        }

        [Test]
        public void BuiltInProperties()
        {
            //ExStart
            //ExFor:PlainTextDocument.BuiltInDocumentProperties
            //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's built-in properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            doc.BuiltInDocumentProperties.Author = "John Doe";

            doc.Save(ArtifactsDir + "PlainTextDocument.BuiltInProperties.docx");

            PlainTextDocument plaintext = new PlainTextDocument(ArtifactsDir + "PlainTextDocument.BuiltInProperties.docx");

            Assert.AreEqual("Hello world!", plaintext.Text.Trim());
            Assert.AreEqual("John Doe", plaintext.BuiltInDocumentProperties.Author);
            //ExEnd
        }

        [Test]
        public void CustomDocumentProperties()
        {
            //ExStart
            //ExFor:PlainTextDocument.CustomDocumentProperties
            //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's custom properties.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            doc.CustomDocumentProperties.Add("Location of writing", "123 Main St, London, UK");

            doc.Save(ArtifactsDir + "PlainTextDocument.CustomDocumentProperties.docx");

            PlainTextDocument plaintext = new PlainTextDocument(ArtifactsDir + "PlainTextDocument.CustomDocumentProperties.docx");

            Assert.AreEqual("Hello world!", plaintext.Text.Trim());
            Assert.AreEqual("123 Main St, London, UK", plaintext.CustomDocumentProperties["Location of writing"].Value);
            //ExEnd
        }
    }
}
