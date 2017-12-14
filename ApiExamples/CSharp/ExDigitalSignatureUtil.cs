// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDigitalSignatureUtil : ApiExampleBase
    {
        [Test]
        public void RemoveAllSignatures()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
            //ExSummary:Shows how to remove every signature from a document.
            //By String:
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outFileName = MyDir + @"\Artifacts\Document.NoSignatures.FromString.doc";

            DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outFileName);

            //By stream:
            Stream streamIn = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            Stream streamOut = new FileStream(MyDir + @"\Artifacts\Document.NoSignatures.FromStream.doc", FileMode.Create);

            DigitalSignatureUtil.RemoveAllSignatures(streamIn, streamOut);
            //ExEnd

            streamIn.Close();
            streamOut.Close();
        }

        [Test]
        public void LoadSignatures()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
            //ExFor:DigitalSignatureUtil.LoadSignatures(String)
            //ExSummary:Shows how to load all existing signatures from a document.
            // By String:
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.LoadSignatures(MyDir + "Document.DigitalSignature.docx");

            // By stream:
            Stream stream = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);

            digitalSignatures = DigitalSignatureUtil.LoadSignatures(stream);
            //ExEnd

            stream.Close();
        }

        [Test]
        public void SignDocument()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, String, DateTime)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, String, DateTime)
            //ExSummary:Shows how to sign documents with personal certificate.
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //By String
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outputFileName = MyDir + @"\Artifacts\Document.DigitalSignature.docx";

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, ch, "My comment", DateTime.Now);

            //By Stream
            Stream streamIn = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            Stream streamOut = new FileStream(MyDir + @"\Artifacts\Document.DigitalSignature.docx", FileMode.OpenOrCreate);

            DigitalSignatureUtil.Sign(streamIn, streamOut, ch, "My comment", DateTime.Now);
            //ExEnd

            streamIn.Dispose();
            streamOut.Dispose();
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
            doc.Save(MyDir + @"\Artifacts\Document.Signed.pdf", options);
            //ExEnd
        }

        //This is for obfuscation bug WORDSNET-13036
        [Test]
        public void SignDocumentTestForBug()
        {
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            Document doc = new Document(MyDir + "TestRepeatingSection.docx");
            String outputFileName = MyDir + @"\Artifacts\TestRepeatingSection.Signed.doc";

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, ch, "My comment", DateTime.Now);
        }

        [Test]
        [ExpectedException(typeof(IncorrectPasswordException), ExpectedMessage = "The document password is incorrect.")]
        public void IncorrectPasswordForDecrypring()
        {
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            Document doc = new Document(MyDir + "Document.Encrypted.docx", new LoadOptions("docPassword"));
            string outputFileName = MyDir + @"\Artifacts\Document.Encrypted.docx";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, ch, "Comment", DateTime.Now, "docPassword1");
        }

        [Test]
        public void SingDocumentWithPasswordDecrypring()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, String, DateTime, String)
            //ExSummary:Shows how to sign encrypted documents.
            // Create certificate holder from a file.
            Document doc = new Document(MyDir + "Document.Encrypted.docx", new LoadOptions("docPassword"));

            string outputFileName = MyDir + @"\Artifacts\Document.Encrypted.docx";

            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, ch, "Comment", DateTime.Now, "docPassword");
            //ExEnd

            // Open encrypted document from a file.
            Document signedDoc = new Document(outputFileName, new LoadOptions("docPassword"));

            // Check that encrypted document was successfully signed.
            DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
            if (signatures.IsValid && (signatures.Count > 0))
            {
                Assert.Pass(); //The document was signed successfully
            }
        }

        [Test]
        public void SingStreamDocumentWithPasswordDecrypring()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, String, DateTime, String)
            //ExSummary:Shows how to sign encrypted documents
            // Create certificate holder from a file.
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            Stream streamIn = new FileStream(MyDir + "Document.Encrypted.docx", FileMode.Open);
            Stream streamOut = new FileStream(MyDir + @"\Artifacts\Document.Encrypted.docx", FileMode.OpenOrCreate);

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(streamIn, streamOut, ch, "Comment", DateTime.Now, "docPassword");
            //ExEnd

            // Open encrypted document from a file.
            Document signedDoc = new Document(streamOut, new LoadOptions("docPassword"));

            // Check that encrypted document was successfully signed.
            DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
            if (signatures.IsValid && (signatures.Count > 0))
            {
                streamIn.Dispose();
                streamOut.Dispose();

                Assert.Pass(); //The document was signed successfully
            }
        }

        [Test]
        public void NoArgumentsForSing()
        {
            Assert.That(() => DigitalSignatureUtil.Sign(String.Empty, String.Empty, null, String.Empty, DateTime.Now, String.Empty), Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void NoCertificateForSign()
        {
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outputFileName = MyDir + @"\Artifacts\Document.DigitalSignature.docx";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            Assert.That(() => DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, null, "Comment", DateTime.Now, "docPassword"), Throws.TypeOf<NullReferenceException>());
        }
    }
}