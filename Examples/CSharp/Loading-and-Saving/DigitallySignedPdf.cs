
using System.IO;
using Aspose.Words;
using System;
//ExStart:X509Certificates
using System.Security.Cryptography.X509Certificates;
//ExEnd:X509Certificates
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class DigitallySignedPdf
    {
        public static void Run()
        {
            //ExStart:DigitallySignedPdf
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Create a simple document from scratch.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Test Signed PDF.");

            // Load the certificate from disk.
            // The other constructor overloads can be used to load certificates from different locations.
            X509Certificate2 cert = new X509Certificate2(
                dataDir + "signature.pfx", "signature");
            

            // Pass the certificate and details to the save options class to sign with.
            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails();
      
            
            dataDir = dataDir + "Document.Signed_out_.pdf";
            // Save the document as PDF.
            doc.Save(dataDir);
            //ExEnd:DigitallySignedPdf
            Console.WriteLine("\nDigitally signed PDF file created successfully.\nFile saved at " + dataDir);
        }
    }
}
