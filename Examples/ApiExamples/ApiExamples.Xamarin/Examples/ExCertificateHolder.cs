using System;
using System.Collections;
using System.IO;
using System.Net;
using System.Security;
using Aspose.Words;
using NUnit.Framework;
#if NET462 || MAC || JAVA
using Org.BouncyCastle.Pkcs;
#endif
namespace ApiExamples
{
    [TestFixture]
    public class ExCertificateHolder : ApiExampleBase
    {
#if NET462 || MAC || JAVA
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
                Pkcs12Store pkcs12Store = new Pkcs12Store(certStream, "aw".ToCharArray());
                IEnumerator enumerator = pkcs12Store.Aliases.GetEnumerator();

                while (enumerator.MoveNext())
                {
                    if (enumerator.Current != null)
                    {
                        string currentAlias = enumerator.Current.ToString();
                        if (pkcs12Store.IsKeyEntry(currentAlias) && pkcs12Store.GetKey(currentAlias).Key.IsPrivate)
                        {
                            Console.WriteLine($"Valid alias found: {enumerator.Current}");
                        }
                    }
                }
            }

            // 3 -  Use a valid alias:
            CertificateHolder.Create(MyDir + "morzal.pfx", "aw", "c20be521-11ea-4976-81ed-865fbbfc9f24");

            // 4 -  Pass "null" as the alias in order to use the first available alias that returns a private key:
            CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);
            //ExEnd
        }
#endif
    }
}