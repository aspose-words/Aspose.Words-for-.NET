// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Drawing;
using NUnit.Framework;
#if NET461_OR_GREATER || JAVA
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExSignDocumentCustom : ApiExampleBase
    {
        //ExStart
        //ExFor:CertificateHolder
        //ExFor:SignatureLineOptions.Signer
        //ExFor:SignatureLineOptions.SignerTitle
        //ExFor:SignatureLine.Id
        //ExFor:SignOptions.SignatureLineId
        //ExFor:SignOptions.SignatureLineImage
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
        //ExSummary:Shows how to add a signature line to a document, and then sign it using a digital certificate.
        [Test] //ExSkip
        [Description("WORDSNET-16868")]
        public static void Sign()
        {
            string signeeName = "Ron Williams";
            string srcDocumentPath = MyDir + "Document.docx";
            string dstDocumentPath = ArtifactsDir + "SignDocumentCustom.Sign.docx";
            string certificatePath = MyDir + "morzal.pfx";
            string certificatePassword = "aw";

            CreateSignees();

            Signee signeeInfo = mSignees.Find(c => c.Name == signeeName);

            if (signeeInfo != null)
                SignDocument(srcDocumentPath, dstDocumentPath, signeeInfo, certificatePath, certificatePassword);
            else
                Assert.Fail("Signee does not exist.");
        }

        /// <summary>
        /// Creates a copy of a source document signed using provided signee information and X509 certificate.
        /// </summary>
        private static void SignDocument(string srcDocumentPath, string dstDocumentPath,
            Signee signeeInfo, string certificatePath, string certificatePassword)
        {
            Document document = new Document(srcDocumentPath);
            DocumentBuilder builder = new DocumentBuilder(document);

            // Configure and insert a signature line, an object in the document that will display a signature that we sign it with.
            SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
            signatureLineOptions.Signer = signeeInfo.Name;
            signatureLineOptions.SignerTitle = signeeInfo.Position;

            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.Id = signeeInfo.PersonId;

            // First, we will save an unsigned version of our document.
            builder.Document.Save(dstDocumentPath);

            CertificateHolder certificateHolder = CertificateHolder.Create(certificatePath, certificatePassword);
            
            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signeeInfo.PersonId;
            signOptions.SignatureLineImage = signeeInfo.Image;

            // Overwrite the unsigned document we saved above with a version signed using the certificate.
            DigitalSignatureUtil.Sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
        }

        public class Signee
        {
            public Guid PersonId { get; set; }
            public string Name { get; set; }
            public string Position { get; set; }
            public byte[] Image { get; set; }

            public Signee(Guid guid, string name, string position, byte[] image)
            {
                PersonId = guid;
                Name = name;
                Position = position;
                Image = image;
            }
        }

        private static void CreateSignees()
        {
            var signImagePath = ImageDir + "Logo.jpg";

            mSignees = new List<Signee>
            {
                new Signee(Guid.NewGuid(), "Ron Williams", "Chief Executive Officer", TestUtil.ImageToByteArray(signImagePath)),
                new Signee(Guid.NewGuid(), "Stephen Morse", "Head of Compliance", TestUtil.ImageToByteArray(signImagePath))
            };
        }

        private static List<Signee> mSignees;
        //ExEnd
    }
}