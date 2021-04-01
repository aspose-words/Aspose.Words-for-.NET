using Aspose.Words;
using System.IO;
using Aspose.Email;
using Aspose.Email.Clients.Smtp;

namespace Convert_a_Document_to_MHTML_and_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"..\..\..\..\Sample Files\";

            // Load a document from the local file system.
            Document doc = new Document(filePath + "MyDocument.docx");

            // Save into a memory stream in MHTML format.
            using (Stream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Mhtml);

                // Rewind the stream to the beginning so Aspose.Email can read it.
                stream.Position = 0;

                // Create an Aspose.Network MIME email message from the stream.
                Aspose.Email.LoadOptions options = new MhtmlLoadOptions();

                MailMessage message = MailMessage.Load(stream, options);
                message.From = "your_from@email.com";
                message.To = "your_to@email.com";
                message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

                // Send the message using Aspose.Email.
                SmtpClient client = new SmtpClient();
                client.Host = "your_smtp.com";
                client.UseAuthentication = false;
                client.Send(message);
            }
        }
    }
}
