// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDigitalSignatureUtil : ApiExampleBase
    {
        [Test]
        public void LoadAndRemove()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil
            //ExFor:DigitalSignatureUtil.LoadSignatures(String)
            //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
            //ExSummary:Shows how to load and remove digital signatures from a digitally signed document.
            // Load digital signatures via filename string to verify that the document is signed
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.LoadSignatures(MyDir + "DigitalSignature.docx");
            Assert.AreEqual(1, digitalSignatures.Count);

            // Re-save the document to an output filename with all digital signatures removed
            DigitalSignatureUtil.RemoveAllSignatures(MyDir + "DigitalSignature.docx", ArtifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

            // Remove all signatures from the document using stream parameters
            using (Stream streamIn = new FileStream(MyDir + "DigitalSignature.docx", FileMode.Open))
            {
                using (Stream streamOut = new FileStream(ArtifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx", FileMode.Create))
                {
                    DigitalSignatureUtil.RemoveAllSignatures(streamIn, streamOut);
                } 
            }

            // We can also load a document's digital signatures via stream, which we will do to verify that all signatures have been removed
            using (Stream stream = new FileStream(ArtifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx", FileMode.Open))
            {
                digitalSignatures = DigitalSignatureUtil.LoadSignatures(stream);
            }

            Assert.AreEqual(0, digitalSignatures.Count);
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
            //ExSummary:Shows how to sign documents using certificate holder and sign options.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            SignOptions signOptions = new SignOptions { Comments = "My comment", SignTime = DateTime.Now };

            using (Stream streamIn = new FileStream(MyDir + "DigitalSignature.docx", FileMode.Open))
            {
                using (Stream streamOut = new FileStream(ArtifactsDir + "DigitalSignatureUtil.SignDocument.docx", FileMode.OpenOrCreate))
                {
                    DigitalSignatureUtil.Sign(streamIn, streamOut, certificateHolder, signOptions);
                }
            }
            //ExEnd
        }

        [Test]
        [Description("WORDSNET-13036, WORDSNET-16868")]
        public void SignDocumentObfuscationBug()
        {
            CertificateHolder ch = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            Document doc = new Document(MyDir + "TestRepeatingSection.docx");
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

            // Digitally sign encrypted with "docPassword" document in the specified path
            Assert.That(
                new TestDelegate(() => DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, certificateHolder, signOptions)),
                Throws.TypeOf<IncorrectPasswordException>(), "The document password is incorrect.");
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
            // Create certificate holder from a file
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            SignOptions signOptions = new SignOptions
            {
                Comments = "Comment",
                SignTime = DateTime.Now,
                DecryptionPassword = "docPassword"
            };

            // Digitally sign encrypted with "docPassword" document in the specified path
            string inputFileName = MyDir + "Encrypted.docx";
            string outputFileName = ArtifactsDir + "DigitalSignatureUtil.DecryptionPassword.docx";

            DigitalSignatureUtil.Sign(inputFileName, outputFileName, certificateHolder, signOptions);
            //ExEnd

            // Open encrypted document from a file
            LoadOptions loadOptions = new LoadOptions("docPassword");
            Assert.AreEqual(signOptions.DecryptionPassword,loadOptions.Password);

            Document signedDoc = new Document(outputFileName, loadOptions);

            // Check that encrypted document was successfully signed
            DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
            if (signatures.IsValid && (signatures.Count > 0))
            {
                //The document was signed successfully
                Assert.Pass();
            }
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
            Document doc = new Document(MyDir + "DigitalSignature.docx");
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