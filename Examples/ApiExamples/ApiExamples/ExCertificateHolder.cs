// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.IO;
using System.Net;
using System.Security;
using Aspose.Words.DigitalSignatures;
using NUnit.Framework;
using Org.BouncyCastle.Pkcs;

namespace ApiExamples
{
    [TestFixture]
    public class ExCertificateHolder : ApiExampleBase
    {
        [Test]
        public void Create()
        {
            //ExStart
            //ExFor:CertificateHolder.Create(Byte[], SecureString)
            //ExFor:CertificateHolder.Create(Byte[], String)
            //ExFor:CertificateHolder.Create(String, String, String)
            //ExSummary:Shows how to create CertificateHolder objects.
            // Below are four ways of creating CertificateHolder objects.
            // 1 -  Load a PKCS #12 file into a byte array and apply its password:
            byte[] certBytes = File.ReadAllBytes(MyDir + "morzal.pfx");
            CertificateHolder.Create(certBytes, "aw");

            // 2 -  Load a PKCS #12 file into a byte array, and apply a secure password:
            SecureString password = new NetworkCredential("", "aw").SecurePassword;
            CertificateHolder.Create(certBytes, password);

            // If the certificate has private keys corresponding to aliases,
            // we can use the aliases to fetch their respective keys. First, we will check for valid aliases.
            using (FileStream certStream = new FileStream(MyDir + "morzal.pfx", FileMode.Open))
            {
                Pkcs12Store pkcs12Store = new Pkcs12StoreBuilder().Build();
                pkcs12Store.Load(certStream, "aw".ToCharArray());
                foreach (string currentAlias in pkcs12Store.Aliases)
                {
                    if ((currentAlias != null) &&
                        (pkcs12Store.IsKeyEntry(currentAlias) &&
                         pkcs12Store.GetKey(currentAlias).Key.IsPrivate))
                    {
                        Console.WriteLine($"Valid alias found: {currentAlias}");
                    }
                }
            }

            // 3 -  Use a valid alias:
            CertificateHolder.Create(MyDir + "morzal.pfx", "aw", "c20be521-11ea-4976-81ed-865fbbfc9f24");

            // 4 -  Pass "null" as the alias in order to use the first available alias that returns a private key:
            CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);
            //ExEnd
        }
    }
}
