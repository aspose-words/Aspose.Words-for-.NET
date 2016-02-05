// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.DigitalSignature
{
    [TestFixture]
    public class ExDigitalSignatureUtil : QaTestsBase
    {
        [Test]
        public void RemoveAllSignaturesEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
            //ExSummary:Shows how to remove every signature from a document.
            //By stream:
            System.IO.Stream docStreamIn = new System.IO.FileStream(ExDir + "Document.Signed.doc", System.IO.FileMode.Open);
            System.IO.Stream docStreamOut = new System.IO.FileStream(ExDir + "Document.NoSignatures.FromStream.doc", System.IO.FileMode.Create);

            DigitalSignatureUtil.RemoveAllSignatures(docStreamIn, docStreamOut);

            docStreamIn.Close();
            docStreamOut.Close();

            //By string:
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.Signed.doc");
            string outFileName = ExDir + "Document.NoSignatures.FromString.doc";

            DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outFileName);
            //ExEnd
        }

        [Test]
        public void LoadSignaturesEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
            //ExFor:DigitalSignatureUtil.LoadSignatures(String)
            //ExSummary:Shows how to load signatures from a document by stream and by string.
            System.IO.Stream docStream = new System.IO.FileStream(ExDir + "Document.Signed.doc", System.IO.FileMode.Open);

            // By stream:
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.LoadSignatures(docStream);
            docStream.Close();

            // By string:
            digitalSignatures = DigitalSignatureUtil.LoadSignatures(ExDir + "Document.Signed.doc");
            //ExEnd
        }

        [Test]
        // We don't include a sample certificate with the examples
        // so this exception is expected instead since the file is not there.
        [ExpectedException(typeof(System.IO.FileNotFoundException))]
        public void SignEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, String, DateTime)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, String, DateTime)
            //ExSummary:Shows how to sign documents.
            CertificateHolder ch = CertificateHolder.Create(ExDir + "MyPkcs12.pfx", "My password");

            //By String
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            string outputDocFileName = ExDir + "Document.Signed.doc";

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "My comment", DateTime.Now);

            //By Stream
            System.IO.Stream docInStream = new System.IO.FileStream(ExDir + "Document.doc", System.IO.FileMode.Open);
            System.IO.Stream docOutStream = new System.IO.FileStream(ExDir + "Document.Signed.doc", System.IO.FileMode.OpenOrCreate);

            DigitalSignatureUtil.Sign(docInStream, docOutStream, ch, "My comment", DateTime.Now);
            //ExEnd
        }
    }
}