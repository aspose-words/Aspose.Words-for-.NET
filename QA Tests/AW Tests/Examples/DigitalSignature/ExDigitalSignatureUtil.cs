// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Document
{
    [TestFixture]
    public class ExDigitalSignatureUtil : QaTestsBase
    {
        [Test]
        public void RemoveAllSignaturesEx()
        {
            //ExStart
            //ExFor:RemoveAllSignatures
            //ExId:RemoveAllSignaturesEx
            //ExSummary:Shows how to use RemoveAllSignatures.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            string outputDocFileName = MyDir + "Document.NoSignatures.doc";
            Aspose.Words.DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outputDocFileName);            
            //ExEnd
        }

        [Test]
        // We don't include a sample certificate with the examples
        // so this exception is expected instead since the file is not there.
        [ExpectedException(typeof(System.IO.FileNotFoundException))]
        public void SignEx()
        {
            //ExStart
            //ExFor:RemoveAllSignatures
            //ExId:RemoveAllSignaturesEx
            //ExSummary:Shows how to use RemoveAllSignatures.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            string outputDocFileName = MyDir + "Document.Signed.doc";

            Aspose.Words.CertificateHolder ch = Aspose.Words.CertificateHolder.Create(MyDir + "MyPkcs12.pfx", "My password");

            Aspose.Words.DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "My comment", DateTime.Now);
            //ExEnd
        }
    }
}
