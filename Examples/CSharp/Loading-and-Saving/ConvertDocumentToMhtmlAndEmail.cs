using Aspose.Words;
using System;
using Aspose.Words.Saving;
using System.IO;
using Aspose.Email;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class ConvertDocumentToMhtmlAndEmail
    {
        public static void Run()
        {
            // ExStart:ConvertDocumentToMhtmlAndEmail
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Load the document into Aspose.Words.
            Document doc = new Document(dataDir + "Test File (docx).docx");

            // Save into a memory stream in MHTML format.
            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);

            // Rewind the stream to the beginning so Aspose.Email can read it.
            stream.Position = 0;

            // Create an Aspose.Network MIME email message from the stream.
            MailMessage message = MailMessage.Load(stream, new MhtmlLoadOptions());
            message.From = "your_from@email.com";
            message.To = "your_to@email.com";
            message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

            // Send the message using Aspose.Email
            SmtpClient client = new SmtpClient();
            client.Host = "your_smtp.com";
            client.Send(message);

            // ExEnd:ConvertDocumentToMhtmlAndEmail

            Console.WriteLine("\nDocument converted to html with roundtrip informations successfully.");
        }
    }
}
