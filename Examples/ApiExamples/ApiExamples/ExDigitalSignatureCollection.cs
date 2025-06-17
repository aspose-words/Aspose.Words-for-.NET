// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using Aspose.Words.DigitalSignatures;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDigitalSignatureCollection : ApiExampleBase
    {
        [Test]
        public void GetEnumerator()
        {
            //ExStart
            //ExFor:DigitalSignatureCollection.GetEnumerator
            //ExSummary:Shows how to print all the digital signatures of a signed document.
            DigitalSignatureCollection digitalSignatures =
                DigitalSignatureUtil.LoadSignatures(MyDir + "Digitally signed.docx");

            using (IEnumerator<DigitalSignature> enumerator = digitalSignatures.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    DigitalSignature ds = enumerator.Current;

                    if (ds != null)
                        Console.WriteLine(ds.ToString());
                }
            }
            //ExEnd

            Assert.That(digitalSignatures.Count, Is.EqualTo(1));

            DigitalSignature signature = digitalSignatures[0];

            Assert.That(signature.IsValid, Is.True);
            Assert.That(signature.SignatureType, Is.EqualTo(DigitalSignatureType.XmlDsig));
            Assert.That(signature.SignTime.ToString("MM/dd/yyyy hh:mm:ss tt"), Is.EqualTo("12/23/2010 02:14:40 AM"));
            Assert.That(signature.Comments, Is.EqualTo("Test Sign"));

            Assert.That(signature.CertificateHolder.Certificate.IssuerName.Name, Is.EqualTo(signature.IssuerName));
            Assert.That(signature.CertificateHolder.Certificate.SubjectName.Name, Is.EqualTo(signature.SubjectName));

            Assert.That(signature.IssuerName, Is.EqualTo("CN=VeriSign Class 3 Code Signing 2009-2 CA, " +
                "OU=Terms of use at https://www.verisign.com/rpa (c)09, " +
                "OU=VeriSign Trust Network, " +
                "O=\"VeriSign, Inc.\", " +
                "C=US"));

            Assert.That(signature.SubjectName, Is.EqualTo("CN=Aspose Pty Ltd, " +
                "OU=Digital ID Class 3 - Microsoft Software Validation v2, " +
                "O=Aspose Pty Ltd, " +
                "L=Lane Cove, " +
                "S=New South Wales, " +
                "C=AU"));
        }
    }
}