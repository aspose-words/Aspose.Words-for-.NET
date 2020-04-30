// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using ApiExamples.TestData.TestClasses;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;
#if NET462 || JAVA
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
#endif

namespace ApiExamples
{
    /// <summary>
    /// This example demonstrates how to add new signature line to the document and sign it with your personal signature <see cref="SignDocument"/>.
    /// </summary>
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
        //ExSummary:Demonstrates how to add new signature line to the document and sign it with personal signature using SignatureLineId.
        [Test] //ExSkip
        [Description("WORDSNET-16868")]
        public static void Sign()
        {
            string signPersonName = "Ron Williams";
            string srcDocumentPath = MyDir + "Document.docx";
            string dstDocumentPath = ArtifactsDir + "SignDocumentCustom.Sign.docx";
            string certificatePath = MyDir + "morzal.pfx";
            string certificatePassword = "aw";

            // We need to create simple list with test signers for this example
            CreateSignPersonData();
            Console.WriteLine("Test data successfully added!");

            // Get sign person object by name of the person who must sign a document
            // This an example, in real use case you would return an object from a database
            SignPersonTestClass signPersonInfo =
                (from c in gSignPersonList where c.Name == signPersonName select c).FirstOrDefault();

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

            // Now do something with a signed document, for example, save it to your database
            // Use 'new Document(dstDocumentPath)' for loading a signed document
        }

        /// <summary>
        /// Signs the document obtained at the source location and saves it to the specified destination.
        /// </summary>
        private static void SignDocument(string srcDocumentPath, string dstDocumentPath,
            SignPersonTestClass signPersonInfo, string certificatePath, string certificatePassword)
        {
            // Create new document instance based on a test file that we need to sign
            Document document = new Document(srcDocumentPath);
            DocumentBuilder builder = new DocumentBuilder(document);

            // Add info about responsible person who sign a document
            SignatureLineOptions signatureLineOptions =
                new SignatureLineOptions { Signer = signPersonInfo.Name, SignerTitle = signPersonInfo.Position };

            // Add signature line for responsible person who sign a document
            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.Id = signPersonInfo.PersonId;

            // Save a document with line signatures into temporary file for future signing
            builder.Document.Save(dstDocumentPath);

            // Create holder of certificate instance based on your personal certificate
            // This is the test certificate generated for this example
            CertificateHolder certificateHolder = CertificateHolder.Create(certificatePath, certificatePassword);

            // Link our signature line with personal signature
            SignOptions signOptions = new SignOptions
            {
                SignatureLineId = signPersonInfo.PersonId,
                SignatureLineImage = signPersonInfo.Image
            };

            // Sign a document which contains signature line with personal certificate
            DigitalSignatureUtil.Sign(dstDocumentPath, dstDocumentPath, certificateHolder, signOptions);
        }

        #if NET462 || JAVA
        /// <summary>
        /// Converting image file to bytes array.
        /// </summary>
        private static byte[] ImageToByteArray(Image imageIn)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                imageIn.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }
        #endif

        /// <summary>
        /// Create test data that contains info about sing persons
        /// </summary>
        private static void CreateSignPersonData()
        {
            gSignPersonList = new List<SignPersonTestClass>
            {
                #if NET462 || JAVA
                new SignPersonTestClass(Guid.NewGuid(), "Ron Williams", "Chief Executive Officer",
                    ImageToByteArray(Image.FromFile(ImageDir + "Logo.jpg"))),
                #elif NETCOREAPP2_1
                new SignPersonTestClass(Guid.NewGuid(), "Ron Williams", "Chief Executive Officer", 
                    SkiaSharp.SKBitmap.Decode(ImageDir + "Logo.jpg").Bytes),
                #endif
                
                #if NET462 || JAVA
                new SignPersonTestClass(Guid.NewGuid(), "Stephen Morse", "Head of Compliance",
                    ImageToByteArray(Image.FromFile(ImageDir + "Logo.jpg")))
                #elif NETCOREAPP2_1
                new SignPersonTestClass(Guid.NewGuid(), "Stephen Morse", "Head of Compliance", 
                    SkiaSharp.SKBitmap.Decode(ImageDir + "Logo.jpg").Bytes)
                #endif
            };
        }

        private static List<SignPersonTestClass> gSignPersonList;
        //ExEnd
    }
}