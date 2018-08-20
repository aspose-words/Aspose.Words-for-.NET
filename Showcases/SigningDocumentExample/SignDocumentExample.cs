// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace SigningDocumentExample
{
    [TestFixture]
    public class SignDocumentExample
    {
        // Sample infrastructure.
        static readonly string ExeDir = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath) + Path.DirectorySeparatorChar;
        static readonly string DataDir = new Uri(new Uri(ExeDir), @"../../Data/").LocalPath;
        static readonly string TestImage = DataDir + @"Images\LogoSmall.png";

        [Test]
        public static void Main(string[] args)
        {
            // We need to create simple List with test signers for this example.
            CreateTestData();
            
            // Get document that we need to sign.
            Document baseDocument = new Document(DataDir + "TestFile.docx");
            DocumentBuilder builder = new DocumentBuilder(baseDocument);

            // Let's define sign person who must sign document.
            string signer = "SignPerson 1";
            
            // Create holder of certificate instanse base on your personal certificate.
            // This is the test certificate generated for this example.
            CertificateHolder certificateHolder = CertificateHolder.Create(DataDir + "morzal.pfx", "aw");

            // Let's add signature to the document and sign it with personal certificate.
            Document signedDocument = SignDocument(signer, builder, certificateHolder);

            // Now we need add signed document to simple List.
            WriteSignedDocument(signedDocument);
        }

        /// <summary>
        /// Add signature line to the document and sign it with personal certificate
        /// </summary>
        /// <param name="signer">Name of the person who sing document</param>
        /// <param name="builder">Class that provides methods for create SignatureLine</param>
        /// <param name="certificateHolder">Holder of personal certificate instanse</param>
        /// <returns>Returns signed document</returns>
        private static Document SignDocument(string signer, DocumentBuilder builder, CertificateHolder certificateHolder)
        {
            // Add info about responsible person who sign document.
            SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
            signatureLineOptions.Signer = GetSignPersonByName(signer).Name;
            signatureLineOptions.SignerTitle = GetSignPersonByName(signer).Position;

            // Add signature line for responsible person who sign document.
            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.Id = GetSignPersonByName(signer).PersonId;

            // Save document with line signatures into temporary file for future signing.
            builder.Document.Save(DataDir + "signedDocument.docx");

            // Link our signature line with personal signature.
            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = GetSignPersonByName(signer).PersonId;
            signOptions.SignatureLineImage = GetSignPersonByName(signer).Image;

            // Sign our document.
            DigitalSignatureUtil.Sign(DataDir + "signedDocument.docx", DataDir + "signedDocument.docx", certificateHolder, signOptions);

            return new Document(DataDir + "signedDocument.docx");
        }

        /// <summary>
        /// Example method for add signed document to simple List
        /// </summary>
        /// <param name="signedDocument">Signed document that we need to add</param>
        private static void WriteSignedDocument(Document signedDocument)
        {
            // This just an example.
            // Actually, it will save or update object with data base.
            mSignDocumentList = new List<SignDocument>
            {
                new SignDocument
                {
                    DocumentId = Guid.NewGuid(),
                    DocumentName = "SignedDocument",
                    Document = ConvertHepler.ConvertDocumentToByteArray(signedDocument)
                }
            };
        }

        /// <summary>
        /// Get sign person object from our simple list
        /// </summary>
        /// <param name="name">Sign person name</param>
        private static SignPerson GetSignPersonByName(string name)
        {
            // This an example.
            // Actually, it will return object from a data base.
            return new Repository<SignPerson>(mSignPersonList.AsQueryable()).FindElement(p => p.Name == name);
        }

        /// <summary>
        /// Add test data to our simple list
        /// </summary>
        private static void CreateTestData()
        {
            mSignPersonList = new List<SignPerson>
            {
                new SignPerson { PersonId = Guid.NewGuid(), Name = "SignPerson 1", Position = "Head of Department", Image = ConvertHepler.ConverImageToByteArray(TestImage) },
                new SignPerson { PersonId = Guid.NewGuid(), Name = "SignPerson 2", Position = "Deputy Head of Department", Image = ConvertHepler.ConverImageToByteArray(TestImage) }
            };
        }

        private static List<SignPerson> mSignPersonList;
        private static List<SignDocument> mSignDocumentList;
    }
}