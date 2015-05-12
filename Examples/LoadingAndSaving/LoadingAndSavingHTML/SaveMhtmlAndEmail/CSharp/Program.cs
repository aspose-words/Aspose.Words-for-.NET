//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;

using Aspose.Words;
#if EmailInstalled
using Aspose.Email.Mail;
#endif

namespace SaveMhtmlAndEmailExample
{
    public class Program
    {
        public static void Main()
        {
#if EmailInstalled

            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //ExStart
            //ExId:SaveMhtmlAndEmail
            //ExSummary:Shows how to save any document from Aspose.Words as MHTML and email using Aspose.Network.
            // Load the document into Aspose.Words.
            string srcFileName = Path.Combine(dataDir, "DinnerInvitationDemo.doc");
            Document doc = new Document(srcFileName);

            // Save into a memory stream in MHTML format.
            Stream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Mhtml);
            // Rewind the stream to the beginning so Aspose.Network can read it.
            stream.Position = 0;

            // Create an Aspose.Network MIME email message from the stream.
            MailMessage message = MailMessage.Load(stream, MessageFormat.Mht);
            message.From = "your_from@email.com";
            message.To = "your_to@email.com";
            message.Subject = "Aspose.Words + Aspose.Network MHTML Test Message";

            // Send the message using Aspose.Network
            SmtpClient client = new SmtpClient();
            client.Host = "your_smtp.com";
            client.AuthenticationMethod = SmtpAuthentication.None;
            client.Send(message);
            //ExEnd
#else
            throw new InvalidOperationException(@"This example requires the use of Aspose.Email." + 
                                                "Make sure Aspose.Email.dll is present in the bin\net2.0 folder.");
#endif
        }
    }
}