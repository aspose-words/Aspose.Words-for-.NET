// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
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
        public void RemoveAllSignatures()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
            //ExSummary:Shows how to remove every signature from a document.
            // By string:
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outFileName = MyDir + @"\Artifacts\Document.NoSignatures.FromString.docx";

            DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outFileName);

            // By stream:
            Stream streamIn = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            Stream streamOut =
                new FileStream(MyDir + @"\Artifacts\Document.NoSignatures.FromStream.doc", FileMode.Create);

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
            // By string:
            DigitalSignatureCollection getDigitalSignaturesByString =
                DigitalSignatureUtil.LoadSignatures(MyDir + "Document.DigitalSignature.docx");

            // By stream:
            Stream stream = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            DigitalSignatureCollection getDigitalSignaturesByStream = DigitalSignatureUtil.LoadSignatures(stream);
            //ExEnd

            stream.Close();
        }

        [Test]
        [Description("WORDSNET-16868")]
        public void SignDocument()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
            //ExFor:SignOptions.Comments
            //ExFor:SignOptions.SignTime
            //ExSummary:Shows how to sign documents using certificate holder and sign options.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            SignOptions signOptions = new SignOptions { Comments = "My comment", SignTime = DateTime.Now };

            Stream streamIn = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            Stream streamOut =
                new FileStream(MyDir + @"\Artifacts\Document.DigitalSignature.docx", FileMode.OpenOrCreate);

            DigitalSignatureUtil.Sign(streamIn, streamOut, certificateHolder, signOptions);
            //ExEnd

            streamIn.Dispose();
            streamOut.Dispose();
        }

        [Test]
        [Description("WORDSNET-13036, WORDSNET-16868")]
        public void SignDocumentObfuscationBug()
        {
            CertificateHolder ch = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            Document doc = new Document(MyDir + "TestRepeatingSection.docx");
            String outputFileName = MyDir + @"\Artifacts\TestRepeatingSection.Signed.doc";

            SignOptions signOptions = new SignOptions { Comments = "Comment", SignTime = DateTime.Now };

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, ch, signOptions);
        }

        [Test]
        [Description("WORDSNET-16868")]
        public void IncorrectPasswordForDecrypring()
        {
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            Document doc = new Document(MyDir + "Document.Encrypted.docx", new LoadOptions("docPassword"));
            string outputFileName = MyDir + @"\Artifacts\Document.Encrypted.docx";

            SignOptions signOptions = new SignOptions
            {
                Comments = "Comment",
                SignTime = DateTime.Now,
                DecryptionPassword = "docPassword1"
            };

            // Digitally sign encrypted with "docPassword" document in the specified path.
            Assert.That(
                () => DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, certificateHolder, signOptions),
                Throws.TypeOf<IncorrectPasswordException>(), "The document password is incorrect.");
        }

        [Test]
        [Description("WORDSNET-16868")]
        public void SingDocumentWithPasswordDecrypring()
        {
            //ExStart
            //ExFor:SignOptions.DecryptionPassword
            //ExSummary:Shows how to sign encrypted document file.
            string outputFileName = MyDir + @"\Artifacts\Document.Encrypted.docx";

            Document doc = new Document(MyDir + "Document.Encrypted.docx", new LoadOptions("docPassword"));

            // Create certificate holder from a file.
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            SignOptions signOptions = new SignOptions
            {
                Comments = "Comment",
                SignTime = DateTime.Now,
                DecryptionPassword = "docPassword"
            };

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputFileName, certificateHolder, signOptions);
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
        public void NoArgumentsForSing()
        {
            SignOptions signOptions = new SignOptions
            {
                Comments = String.Empty,
                SignTime = DateTime.Now,
                DecryptionPassword = String.Empty
            };

            Assert.That(() => DigitalSignatureUtil.Sign(String.Empty, String.Empty, null, signOptions),
                Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void NoCertificateForSign()
        {
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outputFileName = MyDir + @"\Artifacts\Document.DigitalSignature.docx";

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