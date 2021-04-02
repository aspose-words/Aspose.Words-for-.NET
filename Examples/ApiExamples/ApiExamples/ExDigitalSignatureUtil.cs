// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDigitalSignatureUtil : ApiExampleBase
    {
        [Test]
        public void Load()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil
            //ExFor:DigitalSignatureUtil.LoadSignatures(String)
            //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
            //ExSummary:Shows how to load signatures from a digitally signed document.
            // There are two ways of loading a signed document's collection of digital signatures using the DigitalSignatureUtil class.
            // 1 -  Load from a document from a local file system filename:
            DigitalSignatureCollection digitalSignatures = 
                DigitalSignatureUtil.LoadSignatures(MyDir + "Digitally signed.docx");

            // If this collection is nonempty, then we can verify that the document is digitally signed.
            Assert.AreEqual(1, digitalSignatures.Count);

            // 2 -  Load from a document from a FileStream:
            using (Stream stream = new FileStream(MyDir + "Digitally signed.docx", FileMode.Open))
            {
                digitalSignatures = DigitalSignatureUtil.LoadSignatures(stream);
                Assert.AreEqual(1, digitalSignatures.Count);
            }
            //ExEnd
        }

        [Test]
        public void Remove()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil
            //ExFor:DigitalSignatureUtil.LoadSignatures(String)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
            //ExSummary:Shows how to remove digital signatures from a digitally signed document.
            // There are two ways of using the DigitalSignatureUtil class to remove digital signatures
            // from a signed document by saving an unsigned copy of it somewhere else in the local file system.
            // 1 - Determine the locations of both the signed document and the unsigned copy by filename strings:
            DigitalSignatureUtil.RemoveAllSignatures(MyDir + "Digitally signed.docx",
                ArtifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

            // 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
            using (Stream streamIn = new FileStream(MyDir + "Digitally signed.docx", FileMode.Open))
            {
                using (Stream streamOut = new FileStream(ArtifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx", FileMode.Create))
                {
                    DigitalSignatureUtil.RemoveAllSignatures(streamIn, streamOut);
                }
            }

            // Verify that both our output documents have no digital signatures.
            Assert.That(DigitalSignatureUtil.LoadSignatures(ArtifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromString.docx"), Is.Empty);
            Assert.That(DigitalSignatureUtil.LoadSignatures(ArtifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx"), Is.Empty);
            //ExEnd
        }

        [Test]
        [Description("WORDSNET-16868")]
        public void SignDocument()
        {
            //ExStart
            //ExFor:CertificateHolder
            //ExFor:CertificateHolder.Create(String, String)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
            //ExFor:SignOptions.Comments
            //ExFor:SignOptions.SignTime
            //ExSummary:Shows how to digitally sign documents.
            // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            // Create a comment and date which will be applied with our new digital signature.
            SignOptions signOptions = new SignOptions
            {
                Comments = "My comment", 
                SignTime = DateTime.Now
            };

            // Take an unsigned document from the local file system via a file stream,
            // then create a signed copy of it determined by the filename of the output file stream.
            using (Stream streamIn = new FileStream(MyDir + "Document.docx", FileMode.Open))
            {
                using (Stream streamOut = new FileStream(ArtifactsDir + "DigitalSignatureUtil.SignDocument.docx", FileMode.OpenOrCreate))
                {
                    DigitalSignatureUtil.Sign(streamIn, streamOut, certificateHolder, signOptions);
                }
            }
            //ExEnd

            using (Stream stream = new FileStream(ArtifactsDir + "DigitalSignatureUtil.SignDocument.docx", FileMode.Open))
            {
                DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.LoadSignatures(stream);
                Assert.AreEqual(1, digitalSignatures.Count);

                DigitalSignature signature = digitalSignatures[0];

                Assert.True(signature.IsValid);
                Assert.AreEqual(DigitalSignatureType.XmlDsig, signature.SignatureType);
                Assert.AreEqual(signOptions.SignTime.ToString(), signature.SignTime.ToString());
                Assert.AreEqual("My comment", signature.Comments);
            }
        }

        [Test]
        [Description("WORDSNET-16868")]
        public void DecryptionPassword()
        {
            //ExStart
            //ExFor:CertificateHolder
            //ExFor:SignOptions.DecryptionPassword
            //ExFor:LoadOptions.Password
            //ExSummary:Shows how to sign encrypted document file.
            // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            // Create a comment, date, and decryption password which will be applied with our new digital signature.
            SignOptions signOptions = new SignOptions
            {
                Comments = "Comment",
                SignTime = DateTime.Now,
                DecryptionPassword = "docPassword"
            };

            // Set a local system filename for the unsigned input document, and an output filename for its new digitally signed copy.
            string inputFileName = MyDir + "Encrypted.docx";
            string outputFileName = ArtifactsDir + "DigitalSignatureUtil.DecryptionPassword.docx";

            DigitalSignatureUtil.Sign(inputFileName, outputFileName, certificateHolder, signOptions);
            //ExEnd

            // Open encrypted document from a file.
            LoadOptions loadOptions = new LoadOptions("docPassword");
            Assert.AreEqual(signOptions.DecryptionPassword, loadOptions.Password);

            // Check that encrypted document was successfully signed.
            Document signedDoc = new Document(outputFileName, loadOptions);
            DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;

            Assert.AreEqual(1, signatures.Count);
            Assert.True(signatures.IsValid);
        }

        [Test]
        [Description("WORDSNET-13036, WORDSNET-16868")]
        public void SignDocumentObfuscationBug()
        {
            CertificateHolder ch = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            Document doc = new Document(MyDir + "Structured document tags.docx");
            string outputFileName = ArtifactsDir + "DigitalSignatureUtil.SignDocumentObfuscationBug.doc";

            SignOptions signOptions = new SignOptions { Comments = "Comment", SignTime = DateTime.Now };

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, ch, signOptions);
        }

        [Test]
        [Description("WORDSNET-16868")]
        public void IncorrectDecryptionPassword()
        {
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            Document doc = new Document(MyDir + "Encrypted.docx", new LoadOptions("docPassword"));
            string outputFileName = ArtifactsDir + "DigitalSignatureUtil.IncorrectDecryptionPassword.docx";

            SignOptions signOptions = new SignOptions
            {
                Comments = "Comment",
                SignTime = DateTime.Now,
                DecryptionPassword = "docPassword1"
            };

            Assert.That(
                () => DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, certificateHolder, signOptions),
                Throws.TypeOf<IncorrectPasswordException>(), "The document password is incorrect.");
        }

        [Test]
        public void NoArgumentsForSing()
        {
            SignOptions signOptions = new SignOptions
            {
                Comments = string.Empty,
                SignTime = DateTime.Now,
                DecryptionPassword = string.Empty
            };

            Assert.That(() => DigitalSignatureUtil.Sign(string.Empty, string.Empty, null, signOptions),
                Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void NoCertificateForSign()
        {
            Document doc = new Document(MyDir + "Digitally signed.docx");
            string outputFileName = ArtifactsDir + "DigitalSignatureUtil.NoCertificateForSign.docx";

            SignOptions signOptions = new SignOptions
            {
                Comments = "Comment",
                SignTime = DateTime.Now,
                DecryptionPassword = "docPassword"
            };

            Assert.That(() => DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, null, signOptions),
                Throws.TypeOf<ArgumentNullException>());
        }
    }
}