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
            // Load a PKCS #12 file into a byte array and apply its password to create the CertificateHolder
            byte[] certBytes = File.ReadAllBytes(MyDir + "morzal.pfx");
            CertificateHolder.Create(certBytes, "aw");

            // Pass a SecureString which contains the password instead of a normal string
            SecureString password = new NetworkCredential("", "aw").SecurePassword;
            CertificateHolder.Create(certBytes, password);

            // If the certificate has private keys corresponding to aliases, we can use the aliases to fetch their respective keys
            // First, we will check for valid aliases like this
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

            // For this file, we will use an alias found above
            CertificateHolder.Create(MyDir + "morzal.pfx", "aw", "c20be521-11ea-4976-81ed-865fbbfc9f24");

            // If we leave the alias null, then the first possible alias that retrieves a private key will be used
            CertificateHolder.Create(MyDir + "morzal.pfx", "aw", null);
            //ExEnd
        }
#endif
    }
}