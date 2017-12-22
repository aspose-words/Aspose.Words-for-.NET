// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using ApiExamples.TestData;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// This example demonstrates how to add new signature line to the document and sign it with your personal signature <see cref="SignDocument"/>.
    /// </summary>
    [TestFixture]
    public class ExSignDocumentCustom : ApiExampleBase
    {
        [Test]
        //ExStart
        //ExFor:SignatureLineOptions.Signer
        //ExFor:SignatureLineOptions.SignerTitle
        //ExFor:SignatureLine.Id
        //ExFor:SignOptions.SignatureLineId
        //ExFor:SignOptions.SignatureLineImage
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
        //ExSummary:Demonstrates how to add new signature line to the document and sign it with personal signature using SignatureLineId.
        public static void SignSignatureLineUsingSignatureLineId()
        {
            string signPersonName = "Ron Williams";
            string srcDocumentPath = MyDir + "Document.docx";
            string dstDocumentPath = MyDir + @"\Artifacts\Document.Signed.docx";
            string certificatePath = MyDir + "morzal.pfx";
            string certificatePassword = "aw";

            // We need to create simple list with test signers for this example.
            CreateSignPersonData();
            Console.WriteLine("Test data successfully added!");

            // Get sign person object by name of the person who must sign a document.
            // This an example, in real use case you would return an object from a database.
            SignPerson signPersonInfo = (from c in mSignPersonList where c.Name == signPersonName select c).FirstOrDefault();

            if (signPersonInfo != null)
            {
                SignDocument(srcDocumentPath, dstDocumentPath, signPersonInfo, certificatePath, certificatePassword);
                Console.WriteLine("Document successfully signed!");
            }
            else
            {
                Console.WriteLine("Sign person does not exist, please check your parameters.");
                Assert.Fail(); //ExSkip
            }

            // Now do something with a signed document, for example, save it to your database.
            // Use 'new Document(dstDocumentPath)' for loading a signed document.
        }

        /// <summary>
        /// Signs the document obtained at the source location and saves it to the specified destination.
        /// </summary>
        private static void SignDocument(string srcDocumentPath, string dstDocumentPath, SignPerson signPersonInfo, string certificatePath, string certificatePassword)
        {
            // Create new document instance based on a test file that we need to sign.
            Document document = new Document(srcDocumentPath);
            DocumentBuilder builder = new DocumentBuilder(document);

            // Add info about responsible person who sign a document.
            SignatureLineOptions signatureLineOptions = new SignatureLineOptions { Signer = signPersonInfo.Name, SignerTitle = signPersonInfo.Position };

            // Add signature line for responsible person who sign a document.
            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.Id = signPersonInfo.PersonId;

            // Save a document with line signatures into temporary file for future signing.
            builder.Document.Save(dstDocumentPath);

            // Create holder of certificate instance based on your personal certificate.
            // This is the test certificate generated for this example.
            CertificateHolder certificateHolder = CertificateHolder.Create(certificatePath, certificatePassword);

            // Link our signature line with personal signature.
            SignOptions signOptions = new SignOptions { SignatureLineId = signPersonInfo.PersonId, SignatureLineImage = signPersonInfo.Image };

            // Sign a document which contains signature line with personal certificate.
            DigitalSignatureUtil.Sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
        }

        /// <summary>
        /// Converting image file to bytes array
        /// </summary>
        private static byte[] ImageToByteArray(Image imageIn)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                imageIn.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Create test data that contains info about sing persons
        /// </summary>
        private static void CreateSignPersonData()
        {
            mSignPersonList = new List<SignPerson>
            {
                new SignPerson
                {
                    PersonId = Guid.NewGuid(),
                    Name = "Ron Williams",
                    Position = "Chief Executive Officer",
                    Image = ImageToByteArray(Image.FromFile(ImageDir + "LogoSmall.png"))
                },
                new SignPerson
                {
                    PersonId = Guid.NewGuid(),
                    Name = "Stephen Morse",
                    Position = "Head of Compliance",
                    Image = ImageToByteArray(Image.FromFile(ImageDir + "LogoSmall.png"))
                }
            };
        }

        private static List<SignPerson> mSignPersonList;
        //ExEnd
    }
}
