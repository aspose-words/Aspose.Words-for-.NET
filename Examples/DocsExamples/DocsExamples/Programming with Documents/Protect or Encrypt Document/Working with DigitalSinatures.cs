using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Protect_or_Encrypt_Document
{
    internal class WorkingWithDigitalSinatures : DocsExamplesBase
    {
        [Test]
        public void SignDocument()
        {
            //ExStart:SignDocument
            //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
            CertificateHolder certHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            
            DigitalSignatureUtil.Sign(MyDir + "Digitally signed.docx", ArtifactsDir + "Document.Signed.docx",
                certHolder);
            //ExEnd:SignDocument
        }

        [Test]
        public void SigningEncryptedDocument()
        {
            //ExStart:SigningEncryptedDocument
            SignOptions signOptions = new SignOptions { DecryptionPassword = "decryptionPassword" };

            CertificateHolder certHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            
            DigitalSignatureUtil.Sign(MyDir + "Digitally signed.docx", ArtifactsDir + "Document.EncryptedDocument.docx",
                certHolder, signOptions);
            //ExEnd:SigningEncryptedDocument
        }

        [Test]
        public void CreatingAndSigningNewSignatureLine()
        {
            //ExStart:CreatingAndSigningNewSignatureLine
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            SignatureLine signatureLine = builder.InsertSignatureLine(new SignatureLineOptions()).SignatureLine;
            
            doc.Save(ArtifactsDir + "SignDocuments.SignatureLine.docx");

            SignOptions signOptions = new SignOptions
            {
                SignatureLineId = signatureLine.Id,
                SignatureLineImage = File.ReadAllBytes(ImagesDir + "Enhanced Windows MetaFile.emf")
            };

            CertificateHolder certHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            
            DigitalSignatureUtil.Sign(ArtifactsDir + "SignDocuments.SignatureLine.docx",
                ArtifactsDir + "SignDocuments.NewSignatureLine.docx", certHolder, signOptions);
            //ExEnd:CreatingAndSigningNewSignatureLine
        }

        [Test]
        public void SigningExistingSignatureLine()
        {
            //ExStart:SigningExistingSignatureLine
            Document doc = new Document(MyDir + "Signature line.docx");
            
            SignatureLine signatureLine =
                ((Shape) doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true)).SignatureLine;

            SignOptions signOptions = new SignOptions
            {
                SignatureLineId = signatureLine.Id,
                SignatureLineImage = File.ReadAllBytes(ImagesDir + "Enhanced Windows MetaFile.emf")
            };

            CertificateHolder certHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            
            DigitalSignatureUtil.Sign(MyDir + "Digitally signed.docx",
                ArtifactsDir + "SignDocuments.SigningExistingSignatureLine.docx", certHolder, signOptions);
            //ExEnd:SigningExistingSignatureLine
        }

        [Test]
        public void SetSignatureProviderId()
        {
            //ExStart:SetSignatureProviderID
            Document doc = new Document(MyDir + "Signature line.docx");

            SignatureLine signatureLine =
                ((Shape) doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true)).SignatureLine;

            SignOptions signOptions = new SignOptions
            {
                ProviderId = signatureLine.ProviderId, SignatureLineId = signatureLine.Id
            };

            CertificateHolder certHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            DigitalSignatureUtil.Sign(MyDir + "Digitally signed.docx",
                ArtifactsDir + "SignDocuments.SetSignatureProviderId.docx", certHolder, signOptions);
            //ExEnd:SetSignatureProviderID
        }

        [Test]
        public void CreateNewSignatureLineAndSetProviderId()
        {
            //ExStart:CreateNewSignatureLineAndSetProviderId
            //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            SignatureLineOptions signatureLineOptions = new SignatureLineOptions
            {
                Signer = "yourname",
                SignerTitle = "Worker",
                Email = "yourname@aspose.com",
                ShowDate = true,
                DefaultInstructions = false,
                Instructions = "Please sign here.",
                AllowComments = true
            };

            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.ProviderId = Guid.Parse("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2");
            
            doc.Save(ArtifactsDir + "SignDocuments.SignatureLineProviderId.docx");

            SignOptions signOptions = new SignOptions
            {
                SignatureLineId = signatureLine.Id,
                ProviderId = signatureLine.ProviderId,
                Comments = "Document was signed by Aspose",
                SignTime = DateTime.Now
            };

            CertificateHolder certHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            DigitalSignatureUtil.Sign(ArtifactsDir + "SignDocuments.SignatureLineProviderId.docx", 
                ArtifactsDir + "SignDocuments.CreateNewSignatureLineAndSetProviderId.docx", certHolder, signOptions);
            //ExEnd:CreateNewSignatureLineAndSetProviderId
        }

        [Test]
        public void AccessAndVerifySignature()
        {
            //ExStart:AccessAndVerifySignature
            Document doc = new Document(MyDir + "Digitally signed.docx");

            foreach (DigitalSignature signature in doc.DigitalSignatures)
            {
                Console.WriteLine("*** Signature Found ***");
                Console.WriteLine("Is valid: " + signature.IsValid);
                // This property is available in MS Word documents only.
                Console.WriteLine("Reason for signing: " + signature.Comments); 
                Console.WriteLine("Time of signing: " + signature.SignTime);
                Console.WriteLine("Subject name: " + signature.CertificateHolder.Certificate.SubjectName.Name);
                Console.WriteLine("Issuer name: " + signature.CertificateHolder.Certificate.IssuerName.Name);
                Console.WriteLine();
            }
            //ExEnd:AccessAndVerifySignature
        }

        [Test]
        public void RemoveSignatures()
        {
            //ExStart:RemoveSignatures
            //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
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
            //ExEnd:RemoveSignatures
        }
    }
}