using Aspose.Email.Mail;
using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Convert_a_Document_to_MHTML_and_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"E:\Aspose\Aspose Vs VSTO\Aspose.Words Features missing in VSTO 1.1\Sample Files\";
            
            // Load the document into Aspose.Words.
            string srcFileName = MyDir + "Convert_a_Document_to_MHTML_and_Email.doc";
            Document doc = new Document(srcFileName);

            // Save into a memory stream in MHTML format.
            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);

            // Rewind the stream to the beginning so Aspose.Email can read it.
            stream.Position = 0;

            // Create an Aspose.Network MIME email message from the stream.
            MailMessage message = MailMessage.Load(stream, MessageFormat.Mht);
            message.From = "your_from@email.com";
            message.To = "your_to@email.com";
            message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

            // Send the message using Aspose.Email
            SmtpClient client = new SmtpClient();
            client.Host = "your_smtp.com";
            client.AuthenticationMethod = SmtpAuthentication.None;
            client.Send(message);
        }
    }
}
