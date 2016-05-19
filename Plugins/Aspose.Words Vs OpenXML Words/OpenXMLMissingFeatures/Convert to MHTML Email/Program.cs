// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Email.Mail;
using Aspose.Words;
using System.IO;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string FileName = FilePath + "Converting Document.docx";
            // Load the document into Aspose.Words.
            Document doc = new Document(FileName);

            // Save into a memory stream in MHTML format.
            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);

            // Rewind the stream to the beginning so Aspose.Email can read it.
            stream.Position = 0;

           
            // Create an Aspose.Network MIME email message from the stream.
            MailMessage message = MailMessage.Load(stream, new MailMessageLoadOptions(MessageFormat.Mht));
            message.From = "from@gmail.com";
            message.To = "to@gmail.com";
            message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

            // Send the message using Aspose.Email
            SmtpClient client = new SmtpClient();
            client.Host = "smtp.gmail.com";
            client.Port = 587;
            
            client.AuthenticationMethod = SmtpAuthentication.Auto;
            client.Send(message);
        }
    }
}
