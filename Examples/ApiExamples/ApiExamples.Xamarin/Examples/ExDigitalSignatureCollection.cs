// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using Aspose.Words;
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

            Assert.AreEqual(1, digitalSignatures.Count);

            DigitalSignature signature = digitalSignatures[0];

            Assert.True(signature.IsValid);
            Assert.AreEqual(DigitalSignatureType.XmlDsig, signature.SignatureType);
            Assert.AreEqual("12/23/2010 02:14:40 AM", signature.SignTime.ToString("MM/dd/yyyy hh:mm:ss tt"));
            Assert.AreEqual("Test Sign", signature.Comments);

            Assert.AreEqual(signature.IssuerName, signature.CertificateHolder.Certificate.IssuerName.Name);
            Assert.AreEqual(signature.SubjectName, signature.CertificateHolder.Certificate.SubjectName.Name);

            Assert.AreEqual("CN=VeriSign Class 3 Code Signing 2009-2 CA, " +
                "OU=Terms of use at https://www.verisign.com/rpa (c)09, " +
                "OU=VeriSign Trust Network, " +
                "O=\"VeriSign, Inc.\", " +
                "C=US", signature.IssuerName);

            Assert.AreEqual("CN=Aspose Pty Ltd, " +
                "OU=Digital ID Class 3 - Microsoft Software Validation v2, " +
                "O=Aspose Pty Ltd, " +
                "L=Lane Cove, " +
                "S=New South Wales, " +
                "C=AU", signature.SubjectName);
        }
    }
}