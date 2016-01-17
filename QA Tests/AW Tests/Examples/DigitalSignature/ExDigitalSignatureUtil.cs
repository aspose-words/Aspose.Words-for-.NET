// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.DigitalSignature
{
    [TestFixture]
    public class ExDigitalSignatureUtil : QaTestsBase
    {
        public void RemoveAllSignaturesEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(stream, stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(string, string)
            //ExSummary:Shows how to use RemoveAllSignatures.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            //By stream:
            System.IO.Stream docStreamIn = new System.IO.FileStream(ExDir + "Document.doc", System.IO.FileMode.Open);
            System.IO.Stream docStreamOut = new System.IO.FileStream(ExDir + "Document.NoSignatures.FromStream.doc", System.IO.FileMode.Create);

            Aspose.Words.DigitalSignatureUtil.RemoveAllSignatures(docStreamIn, docStreamOut);

            docStreamIn.Close();
            docStreamOut.Close();

            //By string:
            string outFileName = ExDir + "Document.NoSignatures.FromString.doc";

            Aspose.Words.DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outFileName);
            //ExEnd
        }

        [Test]
        public void LoadSignaturesEx()
        {
            //ExStart
            //ExFor:LoadSignatures(stream)
            //ExFor:LoadSignatures(string)
            //ExSummary:Shows how to use LoadSignatures.
            System.IO.Stream docStream = new System.IO.FileStream(ExDir + "Document.doc", System.IO.FileMode.Open);
            Aspose.Words.DigitalSignatureCollection digitalSignatures = new Aspose.Words.DigitalSignatureCollection();

            // By stream:
            digitalSignatures = Aspose.Words.DigitalSignatureUtil.LoadSignatures(docStream);
            docStream.Close();

            // By string:
            digitalSignatures = Aspose.Words.DigitalSignatureUtil.LoadSignatures(ExDir + "Document.doc");
            //ExEnd
        }

        [Test]
        // We don't include a sample certificate with the examples
        // so this exception is expected instead since the file is not there.
        [ExpectedException(typeof(System.IO.FileNotFoundException))]
        public void SignEx()
        {
            //ExStart
            //ExFor:Sign(String, String, CertificateHolder, String, DateTime)
            //ExFor:Sign(Stream, Stream, CertificateHolder, String, DateTime)
            //ExId:SignEx
            //ExSummary:Shows how to use RemoveAllSignatures.
            Aspose.Words.CertificateHolder ch = Aspose.Words.CertificateHolder.Create(ExDir + "MyPkcs12.pfx", "My password");

            //By String
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            string outputDocFileName = ExDir + "Document.Signed.doc";

            Aspose.Words.DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "My comment", DateTime.Now);

            //By Stream
            System.IO.Stream docInStream = new System.IO.FileStream(ExDir + "Document.doc", System.IO.FileMode.Open);
            System.IO.Stream docOutStream = new System.IO.FileStream(ExDir + "Document.Signed.doc", System.IO.FileMode.OpenOrCreate);

            Aspose.Words.DigitalSignatureUtil.Sign(docInStream, docOutStream, ch, "My comment", DateTime.Now);
            //ExEnd
        }
    }
}
