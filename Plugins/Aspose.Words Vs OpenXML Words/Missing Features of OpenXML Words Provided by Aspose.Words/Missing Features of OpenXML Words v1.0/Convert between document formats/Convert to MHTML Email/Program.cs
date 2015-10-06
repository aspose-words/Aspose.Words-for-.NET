// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System.IO;
using Aspose.Email.Mail;
using Aspose.Words;

namespace Convrt_to_MHTML_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            // Load the document into Aspose.Words.
            string srcFileName = Path.Combine(MyDir, "Converting Document.docx");
            Document doc = new Document(srcFileName);

            // Save into a memory stream in MHTML format.
            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);

            // Rewind the stream to the beginning so Aspose.Email can read it.
            stream.Position = 0;

            // Create an Aspose.Network MIME email message from the stream.
            MailMessage message = MailMessage.Load(stream, MessageFormat.Mht);
            message.From = "from@gmail.com";
            message.To = "to@gmail.com";
            message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

            // Send the message using Aspose.Email
            SmtpClient client = new SmtpClient();
            client.Host = "smtp.gmail.com";
            client.Port = 587;
            client.EnableSsl = true;
            client.AuthenticationMethod = SmtpAuthentication.Auto;
            client.Send(message);
        }
    }
}
