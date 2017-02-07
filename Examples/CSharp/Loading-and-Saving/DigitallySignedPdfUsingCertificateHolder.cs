using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class DigitallySignedPdfUsingCertificateHolder
    {
        public static void Run()
        {
            // ExStart:DigitallySignedPdfUsingCertificateHolder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Create a simple document from scratch.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Test Signed PDF.");

            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(
            CertificateHolder.Create(dataDir + "CioSrv1.pfx", "cinD96..arellA"),"reason","location",DateTime.Now);
            doc.Save(dataDir + @"DigitallySignedPdfUsingCertificateHolder.Signed_out.pdf", options);                      
            // ExEnd:DigitallySignedPdfUsingCertificateHolder
            Console.WriteLine("\nDigitally signed PDF file created successfully.\nFile saved at " + dataDir);
        }
    }
}
