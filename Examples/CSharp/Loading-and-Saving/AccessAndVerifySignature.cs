using System;
using System.Collections.Generic;
using Aspose.Words;
using System.IO;

namespace CSharp.Loading_Saving
{
    class AccessAndVerifySignature
    {
        public static void Run()
        {
            //ExStart:AccessAndVerifySignature            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            Document doc = new Document(dataDir + "Test File (doc).doc");
            foreach (DigitalSignature signature in doc.DigitalSignatures)
            {
                Console.WriteLine("*** Signature Found ***");
                Console.WriteLine("Is valid: " + signature.IsValid);
                Console.WriteLine("Reason for signing: " + signature.Comments); // This property is available in MS Word documents only.
                Console.WriteLine("Time of signing: " + signature.SignTime);
                Console.WriteLine("Subject name: " + signature.Certificate.SubjectName.Name);
                Console.WriteLine("Issuer name: " + signature.Certificate.IssuerName.Name);
                Console.WriteLine();
            }
            //ExEnd:AccessAndVerifySignature
        }
    }
}
